#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
#endregion

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreateWallCast : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                //準備東西
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;

                //過濾所有一般模型
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                List<Element> genericModelList = collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel).ToList();
                //確認專案內有放入穿牆開口，並取得其FamilySymbol
                FamilySymbol wallHole = null;
                Autodesk.Revit.DB.View active = doc.ActiveView;
                FilteredElementCollector genericCollector = new FilteredElementCollector(doc, active.Id);
                List<ElementId> genericList = genericCollector.OfCategory(BuiltInCategory.OST_GenericModel).ToElementIds().ToList();
                List<ElementId> holeList = new List<ElementId>();
                foreach (ElementId elemid in genericList)
                {
                    Element elem = doc.GetElement(elemid);
                    ElementId elemtypeid = elem.GetTypeId();
                    ElementType elemtype = doc.GetElement(elemtypeid) as ElementType;
                    if (elemtype.LookupParameter("OmniClass 編號").AsString() == "23.60.30.02.00")
                    {
                        holeList.Add(elemid);
                    }
                }
                foreach (Element el in genericModelList)
                {
                    FamilySymbol elsymbol = el as FamilySymbol;
                    if (elsymbol.LookupParameter("OmniClass 編號").AsString() == "23.60.30.02.00")
                    {
                        wallHole = elsymbol;
                    }
                }
                if (wallHole == null)
                {
                    MessageBox.Show("尚未載入指定的穿牆套管元件");
                    return Result.Failed;
                }
                //拿到管元件
                ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                Reference pickPipeRef = uidoc.Selection.PickObject(ObjectType.Element, pipeFilter, "請選擇貫穿牆的管");
                Element pickPipe = doc.GetElement(pickPipeRef.ElementId);
                //拿到整份牆外參檔&Transform
                ISelectionFilter linkedWallFilter = new WallSelectionFilter(doc);
                Reference pickWallRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedWallFilter, "請選擇貫穿的牆");
                RevitLinkInstance pickWall = doc.GetElement(pickWallRef) as RevitLinkInstance;
                Transform linkTransform = pickWall.GetTotalTransform();

                //拿到實際要用的那道牆元件
                Element linkedWall = pickWall.GetLinkDocument().GetElement(pickWallRef.LinkedElementId);
                Wall wall = linkedWall as Wall;
                double holeLength = UnitUtils.ConvertFromInternalUnits(wall.Width, unitType) + 20;
                LocationCurve wallLocate = linkedWall.Location as LocationCurve;
                Line wallLine = wallLocate.Curve as Line;
                XYZ wallDir = wallLine.Direction;
                XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate();
                XYZ holeDir = XYZ.BasisY;
                double angle = holeDir.AngleTo(wallNorDir);

                ElementId flId = pickPipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                Element fl = doc.GetElement(flId);
                Level level = fl as Level;
                double elevation = level.Elevation;
                double diameter = UnitUtils.ConvertFromInternalUnits(pickPipe.LookupParameter("直徑").AsDouble(), unitType);
                XYZ HoleLocation = GetHoleLocation(linkedWall, pickPipe, linkTransform);
                if (HoleLocation == null)
                {
                    MessageBox.Show("選中的管並未和牆交集");
                }
                FamilySymbol hole = SwitchPipeSize(genericModelList, diameter);
                if (hole == null)
                {
                    message = "專案未載入完整穿牆套管元件或選取管徑≥200mm，請確認!";
                    return Result.Failed;
                }
                //bool isintersect = isOccupy(HoleLocation, doc, holeList);
                //if (isintersect)
                double cl = HoleLocation.Z - elevation;
                FamilyInstance createHole = null;
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("創建套管");
                    hole.Activate();
                    createHole = doc.Create.NewFamilyInstance(HoleLocation, hole, fl, level, StructuralType.NonStructural);
                    createHole.LookupParameter("開口CL垂直偏移").Set(cl);
                    createHole.LookupParameter("L").Set(UnitUtils.ConvertToInternalUnits(holeLength, unitType));
                    createHole.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(0);
                    LocationPoint loPoint = createHole.Location as LocationPoint;
                    XYZ newPoint = new XYZ(HoleLocation.X, HoleLocation.Y, HoleLocation.Z + 10);
                    Line axis = Line.CreateBound(HoleLocation, newPoint);
                    loPoint.Rotate(axis, angle);
                    trans.Commit();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }

            return Result.Succeeded;
        }
        //取得管道與牆干涉的點(XYZ)
        public XYZ GetHoleLocation(Element wallElem, Element pipeElem, Transform trans)
        {
            //取得牆的solid
            Solid solid_wall = singleSolidFromElement(wallElem);
            solid_wall = SolidUtils.CreateTransformed(solid_wall, trans);
            //取得管的curve
            LocationCurve pipe_locationcurve = pipeElem.Location as LocationCurve;
            Curve pipe_Line = pipe_locationcurve.Curve as Curve;
            //取得管與牆相交的curve
            SolidCurveIntersectionOptions intersectOptions = new SolidCurveIntersectionOptions();
            SolidCurveIntersection intersection = solid_wall.IntersectWithCurve(pipe_Line, intersectOptions);
            XYZ point_Center = null;
            if (intersection.SegmentCount > 0)
            {
                Curve curveInside = intersection.GetCurveSegment(0);
                //取得curve中點
                XYZ point_Start = curveInside.GetEndPoint(0);
                XYZ point_End = curveInside.GetEndPoint(1);
                point_Center = new XYZ((point_Start.X + point_End.X) / 2, (point_Start.Y + point_End.Y) / 2, ((point_Start.Z + point_End.Z) / 2));
            }
            return point_Center;
        }

        //element轉換為solid
        public Solid singleSolidFromElement(Element element)
        {
            Options options = new Options();
            GeometryElement geomElem = element.get_Geometry(options);
            Solid solidResult = null;
            foreach (GeometryObject geomObj in geomElem)
            {
                solidResult = geomObj as Solid;
            }
            return solidResult;
        }
        //判斷放置穿牆開口的類型
        public FamilySymbol SwitchPipeSize(List<Element> genericModelList, double width)
        {
            FamilySymbol hole = null;
            FamilySymbol d40 = null; FamilySymbol d50 = null; FamilySymbol d80 = null; FamilySymbol d100 = null;
            FamilySymbol d125 = null; FamilySymbol d150 = null; FamilySymbol d200 = null; FamilySymbol d250 = null;
            foreach (Element el in genericModelList)
            {
                FamilySymbol elsymbol = el as FamilySymbol;
                if (elsymbol.LookupParameter("OmniClass 編號").AsString() == "23.60.30.02.00")
                {
                    if (elsymbol.LookupParameter("D") != null)
                    {
                        switch (elsymbol.LookupParameter("D").AsValueString())
                        {
                            case "42":
                                d40 = elsymbol;
                                break;
                            case "52":
                                d50 = elsymbol;
                                break;
                            case "82":
                                d80 = elsymbol;
                                break;
                            case "102":
                                d100 = elsymbol;
                                break;
                            case "127":
                                d125 = elsymbol;
                                break;
                            case "152":
                                d150 = elsymbol;
                                break;
                            case "202":
                                d200 = elsymbol;
                                break;
                            case "252":
                                d250 = elsymbol;
                                break;
                        }
                    }
                }
            }
            if (width < 20)
            {
                hole = d50;
            }
            else if (20 <= width && width < 40)
            {
                hole = d50;
            }
            else if (40 <= width && width < 65)
            {
                hole = d80;
            }
            else if (65 <= width && width < 75)
            {
                hole = d100;
            }
            else if (75 <= width && width < 100)
            {
                hole = d125;
            }
            else if (100 <= width && width < 125)
            {
                hole = d150;
            }
            else if (125 <= width && width < 150)
            {
                hole = d200;
            }
            else if (150 <= width && width < 200)
            {
                hole = d250;
            }
            return hole;
        }
        //XYZ判斷該處是否已有元件
        public bool isOccupy(XYZ point, Document doc, List<ElementId> holeList)
        {
            bool isOccupy = true;
            if (holeList.Count() > 0)
            {
                XYZ max = new XYZ(point.X + .05, point.Y + .05, point.Z + .05);
                XYZ min = new XYZ(point.X - .05, point.Y - .05, point.Z - .05);
                Outline box = new Outline(min, max);
                FilteredElementCollector collector = new FilteredElementCollector(doc, holeList);
                BoundingBoxIntersectsFilter boxFilter = new BoundingBoxIntersectsFilter(box);
                List<Element> boxIntersect = collector.WherePasses(boxFilter).OfCategory(BuiltInCategory.OST_GenericModel).ToList();
                if (boxIntersect.Count() == 0)
                {
                    isOccupy = false;
                }
            }
            else
            {
                isOccupy = false;
            }
            return isOccupy;
        }
        public static XYZ TransformPoint(XYZ point, Transform transform)
        {
            double x = point.X;
            double y = point.Y;
            double z = point.Z;
            XYZ val = transform.get_Basis(0);
            XYZ val2 = transform.get_Basis(1);
            XYZ val3 = transform.get_Basis(2);
            XYZ origin = transform.Origin;
            double xTemp = x * val.X + y * val2.X + z * val3.X + origin.X;
            double yTemp = x * val.Y + y * val2.Y + z * val3.Y + origin.Y;
            double zTemp = x * val.Z + y * val2.Z + z * val3.Z + origin.Z;
            return new XYZ(xTemp, yTemp, zTemp);
        }
    }
}
//品類過濾器

