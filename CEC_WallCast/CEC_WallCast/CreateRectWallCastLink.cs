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
    class CreateRectWallCastLink : IExternalCommand
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

                    //拿到外參管元件
                    ISelectionFilter linkpipeFilter = new linkedPipeSelectionFilter(doc);
                    Reference refer = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkpipeFilter, "請選擇貫穿牆的「連結模型」管");
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTrans = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);

                    //拿到整份外參牆&Transform
                    ISelectionFilter linkedWallFilter = new WallSelectionFilter(doc);
                    Reference pickWallRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedWallFilter, "請選擇「連結模型」中被貫穿的牆");
                    RevitLinkInstance pickWall = doc.GetElement(pickWallRef) as RevitLinkInstance;
                    Transform linkTransform = pickWall.GetTotalTransform();

                    Element linkedWall = pickWall.GetLinkDocument().GetElement(pickWallRef.LinkedElementId);
                    Wall wall = linkedWall as Wall;
                    double holeLength = UnitUtils.ConvertFromInternalUnits(wall.Width, unitType) + 20;
                    LocationCurve wallLocate = linkedWall.Location as LocationCurve;
                    Curve wallCrv = wallLocate.Curve;
                    wallCrv = wallCrv.CreateTransformed(linkTransform);
                    Line wallLine = wallCrv as Line;
                    double angle = 0.0;
                    bool isLiner = false;
                    XYZ holeDir = XYZ.BasisY;
                    if (wallLine != null)
                    {
                        isLiner = true;
                        XYZ wallDir = wallLine.Direction.Normalize();
                        XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate(); //這段要再看一下
                        angle = holeDir.AngleTo(wallNorDir);
                    }
                    //1.擷取每支管和這道牆的交界點
                    MEPCurve linkPipeCrv = linkedPipe as MEPCurve;
                    LocationCurve pipeLocate = linkPipeCrv.Location as LocationCurve;
                    Curve pipeCurve = pipeLocate.Curve;
                    //Level pipeLevel = doc.GetElement(linkPipeCrv.ReferenceLevel.Id) as Level;

                    Level pipeLevel = null;
                    string LevelName = linkPipeCrv.ReferenceLevel.Name;
                    Level sourceLevel = linkPipeCrv.ReferenceLevel;
                    FilteredElementCollector levelFilter = new FilteredElementCollector(doc).OfClass(typeof(Level));
                    foreach (Level le in levelFilter)
                    {
                        if (le.Name == LevelName || le.ProjectElevation == sourceLevel.ProjectElevation)
                        {
                            pipeLevel = le;
                        }
                    }

                    if (pipeLevel == null)
                    {
                        MessageBox.Show("請確認「連結模型」中的樓層命名原則是否和本機端一致");
                        return Result.Failed;
                    }
                    double elevation = pipeLevel.ProjectElevation;
                    XYZ HoleLocation = GetHoleLocation(linkedWall, linkedPipe, linkTransform);
                    if (HoleLocation == null)
                    {
                        MessageBox.Show("執行失敗，請確認選中的管是否都有和牆交集!");
                        return Result.Failed;
                    }
                    //如果牆不為曲線，須單獨設定，計算曲線導數作為旋轉依據
                    if (isLiner != true)
                    {
                        IntersectionResult intersect = wallCrv.Project(HoleLocation);
                        Transform ptTrans = wallCrv.ComputeDerivatives(intersect.Parameter, false);
                        XYZ v = ptTrans.BasisY;
                        angle = holeDir.AngleTo(v);
                    }
                    double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                    Family Wall_Cast;
                    FamilyInstance instance = null;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        Wall_Cast = new RectWallCast().WallCastSymbol(doc);
                        tx.Commit();
                    }
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置穿牆套管");
                        if (HoleLocation != null)
                        {
                            FamilySymbol CastSymbol2 = new RectWallCast().findWall_CastSymbol(doc, Wall_Cast, linkedPipe);
                            Parameter pipeWidth = getPipeWidth(linkedPipe);
                            Parameter pipeHigh = getPipeHeight(linkedPipe);
                            instance = doc.Create.NewFamilyInstance(HoleLocation, CastSymbol2, pipeLevel, StructuralType.NonStructural);
                            //參數檢查
                            List<string> paraNameToCheck = new List<string>()
                                {
                                   "開口長","系統別","TTOP","TCOP","TBOP","BBOP","BCOP","BTOP"
                                };

                            foreach (string item in paraNameToCheck)
                            {
                                if (!checkPara(instance, item))
                                {
                                    MessageBox.Show($"執行失敗，請檢查 {instance.Symbol.FamilyName} 元件中是否缺少 {item} 參數欄位");
                                    return Result.Failed;
                                }
                            }

                            //2022.05.13_新增：針對具有斜率的管材，計算對應的偏移值
                            double slopeOffset = 0.0;
                            Parameter pipeSlope = linkedPipe.LookupParameter("斜度");
                            if (pipeSlope != null && pipeSlope.AsDouble() != 0)
                            {
                                double pipeStartHeight = linkedPipe.get_Parameter(BuiltInParameter.RBS_START_OFFSET_PARAM).AsDouble();
                                double pipeEndHeight = linkedPipe.get_Parameter(BuiltInParameter.RBS_END_OFFSET_PARAM).AsDouble();
                                XYZ startPt = pipeCurve.GetEndPoint(0);
                                XYZ endPt = pipeCurve.GetEndPoint(1);
                                double distToStart = HoleLocation.DistanceTo(startPt);
                                if (pipeStartHeight >= pipeEndHeight)
                                {
                                    slopeOffset = -(distToStart * pipeSlope.AsDouble());
                                }
                                else if (pipeStartHeight < pipeEndHeight)
                                {
                                    slopeOffset = distToStart * pipeSlope.AsDouble();
                                }
                            }
                            double castOffset = UnitUtils.ConvertToInternalUnits(100, unitType);
                            double heightToSet = pipeHigh.AsDouble() + castOffset;
                            double widthToSet = pipeWidth.AsDouble() + castOffset;
                            instance.LookupParameter("開口高").Set(heightToSet);
                            instance.LookupParameter("開口寬").Set(widthToSet);


                            //調整高度與長度
                            double castDiameter = heightToSet / 2;
                            //double castDiameter = instance.LookupParameter("開口高").AsDouble() / 4;
                            //MessageBox.Show(UnitUtils.ConvertFromInternalUnits(heightToSet, unitType).ToString());
                            //MessageBox.Show(UnitUtils.ConvertFromInternalUnits(castDiameter, unitType).ToString());
                            //double pipeHeight = pickPipe.LookupParameter("偏移").AsDouble();
                            double pipeHeight = linkedPipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble();
                            instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(pipeHeight - castDiameter + slopeOffset);
                            instance.LookupParameter("開口長").Set(castLength);
                            LocationPoint loPoint = instance.Location as LocationPoint;
                            XYZ newPoint = new XYZ(HoleLocation.X, HoleLocation.Y, HoleLocation.Z + 10);
                            Line axis = Line.CreateBound(HoleLocation, newPoint);
                            loPoint.Rotate(axis, angle);


                            //設定BBOP、BCOP、BTOP (牆只需要設定從底部開始的參數)
                            Parameter BBOP = instance.LookupParameter("BBOP");
                            Parameter BCOP = instance.LookupParameter("BCOP");
                            Parameter BTOP = instance.LookupParameter("BTOP");
                            Parameter TTOP = instance.LookupParameter("TTOP");
                            Parameter TCOP = instance.LookupParameter("TCOP");
                            Parameter TBOP = instance.LookupParameter("TBOP");
                            double outterDiameter = instance.LookupParameter("開口高").AsDouble();
                            double wallBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                            double wallHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                            //設定值計算，後來修正成只要用FL計算即可
                            //抓到模型中所有的樓層元素，依照樓高排序。要找到位於他上方的樓層
                            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                            ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                            levelCollector.WherePasses(level_Filter).WhereElementIsNotElementType().ToElements();
                            List<string> levelNames = new List<string>(); //用名字來確認篩選排序
                            List<Element> level_List = levelCollector.OrderBy(x => sortLevelbyHeight(x)).ToList();
                            for (int i = 0; i < level_List.Count(); i++)
                            {
                                Level le = level_List[i] as Level;
                                levelNames.Add(le.Name);
                            }
                            int index_lowLevel = levelNames.IndexOf(pipeLevel.Name);
                            int index_topLevel = index_lowLevel + 1;
                            Level topLevel = null;
                            if (index_topLevel < level_List.Count())
                            {
                                topLevel = level_List[index_topLevel] as Level;
                            }
                            else if (topLevel == null)
                            {
                                message = "管的上方沒有樓層，無法計算穿牆套管偏移值";
                                return Result.Failed;
                            }
                            double basicWallHeight = topLevel.Elevation - pipeLevel.Elevation;
                            double BBOP_toSet = pipeHeight - outterDiameter / 2;
                            double BCOP_toSet = pipeHeight;
                            double BTOP_toSet = pipeHeight + outterDiameter / 2;
                            double TCOP_toSet = basicWallHeight - BCOP_toSet;
                            double TBOP_toSet = TCOP_toSet + outterDiameter / 2;
                            double TTOP_toSet = TCOP_toSet - outterDiameter / 2;
                            BBOP.Set(BBOP_toSet);
                            BCOP.Set(BCOP_toSet);
                            BTOP.Set(BTOP_toSet);
                            TBOP.Set(TBOP_toSet);
                            TCOP.Set(TCOP_toSet);
                            TTOP.Set(TTOP_toSet);
                        }
                        trans.Commit();
                    }
                }
                catch
                {
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
        public double sortLevelbyHeight(Element element)
        {
            Level tempLevel = element as Level;
            double levelHeight = element.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();
            //double levelHeight = element.LookupParameter("立面").AsDouble();
            return levelHeight;
        }
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
        public Parameter getPipeWidth(Element element)
        {
            Parameter targetPara = null;
            //Pipe >用外徑計算
            if (element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            }
            //Conduit
            else if (element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
            }
            //Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            }
            //方型Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            }
            //電纜架
            else if (element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            }
            return targetPara;
        }
        public Parameter getPipeHeight(Element element)
        {
            Parameter targetPara = null;
            //Pipe
            if (element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            }
            //Conduit
            else if (element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
            }
            //Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            }
            //方型Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            }
            //電纜架
            else if (element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
            }
            return targetPara;
        }
    }
}
