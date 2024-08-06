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
    class CreateRectSlabCast : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            while (true)
            {
                Counter.count += 1;
                try
                {
                    //準備東西
                    UIApplication uiapp = commandData.Application;
                    UIDocument uidoc = uiapp.ActiveUIDocument;
                    Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                    Document doc = uidoc.Document;
                    method m = new method();

                    //拿到管元件
                    ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                    Reference pickPipeRef = uidoc.Selection.PickObject(ObjectType.Element, pipeFilter, "請選擇貫穿版的管");
                    Element pickPipe = doc.GetElement(pickPipeRef.ElementId);
                    //拿到整份牆外參檔&Transform
                    ISelectionFilter linkedWallFilter = new WallSelectionFilter(doc);
                    Reference pickSlabRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedWallFilter, "請選擇貫穿的版");
                    RevitLinkInstance pickSlab = doc.GetElement(pickSlabRef) as RevitLinkInstance;
                    Transform linkTransform = pickSlab.GetTotalTransform();

                    //拿到實際要用的那道版元件
                    Document linkDoc = pickSlab.GetLinkDocument();
                    Element linkedSlab = linkDoc.GetElement(pickSlabRef.LinkedElementId);
                    Level linkLevel = linkDoc.GetElement(linkedSlab.LevelId) as Level;
                    double linkLevelElevation = linkLevel.ProjectElevation;
                    Level targetLevel = null;
                    FilteredElementCollector levelColl = new FilteredElementCollector(doc).OfClass(typeof(Level)).WhereElementIsNotElementType();
                    foreach (Element e in levelColl)
                    {
                        Level tempLevel = e as Level;
                        if (tempLevel.ProjectElevation == linkLevelElevation)
                        {
                            targetLevel = tempLevel;
                            break;
                        }
                    }

                    Floor slab = linkedSlab as Floor;
                    Solid slabSolid = m.singleSolidFromElement(linkedSlab);
                    slabSolid = SolidUtils.CreateTransformed(slabSolid, linkTransform);
                    double slabWidth = linkedSlab.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble();
                    double holeLength = UnitUtils.ConvertFromInternalUnits(slabWidth, unitType) + 20;

                    MEPCurve pipeCrv = pickPipe as MEPCurve;
                    LocationCurve pipeLocate = pipeCrv.Location as LocationCurve;
                    Curve pipeCurve = pipeLocate.Curve;
                    Solid tempPipeSolid = m.singleSolidFromElement(pickPipe);
                    Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(slabSolid, tempPipeSolid, BooleanOperationsType.Intersect);
                    XYZ targetPt = null;
                    double holeWidth = 0.0;
                    double holeHeight = 0.0;
                    if (Math.Abs(interSolid.Volume) > 0.000001)
                    {
                        targetPt = interSolid.GetBoundingBox().Transform.Origin;
                        holeWidth =interSolid.GetBoundingBox().Max.X - interSolid.GetBoundingBox().Min.X;
                        holeHeight= interSolid.GetBoundingBox().Max.Y - interSolid.GetBoundingBox().Min.Y;
                    }
                    else
                    {
                        MessageBox.Show("管沒有和任何的版交集，請重新調整!");
                    }

                    //Level level = pipeCrv.ReferenceLevel; //取得管線的參考樓層
                    //double elevation = level.Elevation;
                    //XYZ HoleLocation = GetHoleLocation(linkedSlab, pickPipe, linkTransform);
                    //if (HoleLocation == null)
                    //{
                    //    MessageBox.Show("管沒有和任何的版交集，請重新調整!");
                    //}
                    Family Slab_Cast;
                    FamilyInstance instance = null;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        Slab_Cast = new RectSlabCast().SlabCastSymbol(doc);
                        tx.Commit();
                    }
         
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置穿版套管");
                        if (targetPt != null)
                        {
                            FamilySymbol CastSymbol2 = new RectSlabCast().findSlab_CastSymbol(doc, Slab_Cast);
                            //Parameter pipeWidth = m.getPipeWidth(pickPipe);
                            //Parameter pipeHeight = m.getPipeHeight(pickPipe);
                            instance = doc.Create.NewFamilyInstance(targetPt, CastSymbol2, targetLevel, StructuralType.NonStructural);
                            //參數檢查
                            List<string> paraNameToCheck = new List<string>()
                                {
                                   "開口長","開口高","開口長","系統別"
                                };

                            foreach (string item in paraNameToCheck)
                            {
                                if (!checkPara(instance, item))
                                {
                                    MessageBox.Show($"執行失敗，請檢查 {instance.Symbol.FamilyName} 元件中是否缺少 {item} 參數欄位");
                                    return Result.Failed;
                                }
                            }

#if RELEASE2019
                            Parameter instHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
#else
                            Parameter instHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
#endif

                            double castOffset = UnitUtils.ConvertToInternalUnits(100, unitType);
                            double heightToSet = holeHeight + castOffset;
                            double widthToSet = holeWidth + castOffset;
                            instance.LookupParameter("開口高").Set(heightToSet);
                            instance.LookupParameter("開口寬").Set(widthToSet);

                            instHeight.Set(targetPt.Z - targetLevel.ProjectElevation + slabWidth / 2);
                            double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                            instance.LookupParameter("開口長").Set(castLength);
                        }
                        trans.Commit();
                    }
                }
                catch
                {
                    //MessageBox.Show("執行失敗");
                    //return Result.Failed;
                    break;
                }
            }
            return Result.Succeeded;
        }
        public bool checkPara(Element elem, string paraName)
        {
            bool result = false;
            foreach (Parameter parameter in elem.Parameters)
            {
                Parameter val = parameter;
                if (val.Definition.Name == paraName)
                {
                    result = true;
                }
            }
            return result;
        }
        public static IList<Solid> GetTargetSolids(Element element)
        {
            List<Solid> solids = new List<Solid>();
            Options options = new Options();
            //預設為不包含不可見元件，因此改成true
            options.ComputeReferences = true;
            options.DetailLevel = ViewDetailLevel.Fine;
            options.IncludeNonVisibleObjects = true;
            GeometryElement geomElem = element.get_Geometry(options);
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid)
                {
                    Solid solid = (Solid)geomObj;
                    if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                    {
                        solids.Add(solid);
                    }
                }
                else if (geomObj is GeometryInstance)//一些特殊狀況可能會用到，like樓梯
                {
                    GeometryInstance geomInst = (GeometryInstance)geomObj;
                    GeometryElement instGeomElem = geomInst.GetInstanceGeometry();
                    foreach (GeometryObject instGeomObj in instGeomElem)
                    {
                        if (instGeomObj is Solid)
                        {
                            Solid solid = (Solid)instGeomObj;
                            if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                            {
                                solids.Add(solid);
                            }
                        }
                    }
                }
            }
            return solids;
        }
        public static Solid singleSolidFromElement(Element inputElement)
        {
            Document doc = inputElement.Document;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            // create solid from Element:
            IList<Solid> fromElement = GetTargetSolids(inputElement);
            int solidCount = fromElement.Count;
            // MessageBox.Show(solidCount.ToString());
            // Merge all found solids into single one
            Solid solidResult = null;
            //XYZ checkheight = new XYZ(0, 0, 6.88976);
            //Transform tr = Transform.CreateTranslation(checkheight);
            if (solidCount == 1)
            {
                solidResult = fromElement[0];
            }
            else if (solidCount > 1)
            {
                solidResult =
                    BooleanOperationsUtils.ExecuteBooleanOperation(fromElement[0], fromElement[1], BooleanOperationsType.Union);
            }

            if (solidCount > 2)
            {
                for (int i = 2; i < solidCount; i++)
                {
                    solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(solidResult, fromElement[i], BooleanOperationsType.Union);
                }
            }
            return solidResult;
        }
        public XYZ GetHoleLocation(Element wallElem, Element pipeElem, Transform trans)
        {

            //取得牆的solid
            Solid solid_wall = singleSolidFromElement(wallElem);
            solid_wall = SolidUtils.CreateTransformed(solid_wall, trans);
            //取得管的curve
            LocationCurve pipe_locationcurve = pipeElem.Location as LocationCurve;
            Curve pipe_Line = pipe_locationcurve.Curve;
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
                XYZ tempCenter = curveInside.Evaluate(0.5, true);
                //point_Center = new XYZ((point_Start.X + point_End.X) / 2, (point_Start.Y + point_End.Y) / 2, ((point_Start.Z + point_End.Z) / 2));
                point_Center = tempCenter;

            }
            return point_Center;

        }
    }
    public class RectSlabCast
    {
        #region
        //將穿版套管的功能做成class管理
        //1.先找到套管的Family
        //2.用Family反查Symbol
        #endregion
        public Family SlabCastSymbol(Document doc)
        {
            string internalNameWall = "CEC-穿版套管-方";
            Family Wall_CastType = null;
            ElementFilter Wall_CastCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
            ElementFilter Wall_CastSymbolFilter = new ElementClassFilter(typeof(FamilySymbol));

            LogicalAndFilter andFilter = new LogicalAndFilter(Wall_CastCategoryFilter, Wall_CastSymbolFilter);
            FilteredElementCollector RC_CastSymbol = new FilteredElementCollector(doc);
            RC_CastSymbol = RC_CastSymbol.WherePasses(andFilter);//這地方有點怪，無法使用andFilter RC_CastSymbolFilter
            bool symbolFound = false;
            foreach (FamilySymbol e in RC_CastSymbol)
            {
                Parameter p = e.LookupParameter("API識別名稱");
                if (p != null && p.AsString().Contains(internalNameWall))
                {
                    symbolFound = true;
                    Wall_CastType = e.Family;
                    break;
                }
            }
            if (!symbolFound)
            {
                MessageBox.Show("尚未載入指定的穿版套管元件!");
            }
            return Wall_CastType;
        }
        public FamilySymbol findSlab_CastSymbol(Document doc, Family CastFamily)
        {
            FamilySymbol targetFamilySymbol = null;
            if (CastFamily != null)
            {
                FamilySymbol tempSymbol = doc.GetElement(CastFamily.GetFamilySymbolIds().First()) as FamilySymbol;
                targetFamilySymbol = tempSymbol;
            }
            targetFamilySymbol.Activate();
            return targetFamilySymbol;
        }

    }
}
