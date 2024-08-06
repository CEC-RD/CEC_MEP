using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreateRectWallCast : IExternalCommand
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
                    Curve wallCrv = wallLocate.Curve;
                    wallCrv = wallCrv.CreateTransformed(linkTransform);
                    Line wallLine = wallCrv as Line;
                    double angle = 0.0;
                    bool isLiner = false;
                    XYZ holeDir = XYZ.BasisY;
                    if (wallLine != null)
                    {
                        isLiner = true;
                        XYZ wallDir = wallLine.Direction;
                        XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate();
                        angle = holeDir.AngleTo(wallNorDir);
                    }

                    MEPCurve pipeCrv = pickPipe as MEPCurve;
                    LocationCurve pipeLocate = pipeCrv.Location as LocationCurve;
                    Curve pipeCurve = pipeLocate.Curve;
                    Level level = pipeCrv.ReferenceLevel; //取得管線的參考樓層
                    double elevation = level.Elevation;
                    XYZ HoleLocation = GetHoleLocation(linkedWall, pickPipe, linkTransform);
                    if (HoleLocation == null)
                    {
                        MessageBox.Show("管沒有和任何的牆交集，請重新調整!");
                    }
                    //如果牆不為曲線，須單獨設定，計算曲線導數作為旋轉依據
                    if (isLiner != true)
                    {
                        IntersectionResult intersect = wallCrv.Project(HoleLocation);
                        Transform ptTrans = wallCrv.ComputeDerivatives(intersect.Parameter, false);
                        XYZ v = ptTrans.BasisY;
                        angle = holeDir.AngleTo(v);
                    }

                    Family Wall_Cast;
                    FamilyInstance instance = null;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        Wall_Cast = new RectWallCast().WallCastSymbol(doc);
                        tx.Commit();
                    }
                    #region 放置穿牆套管開始
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置穿牆套管");
                        if (HoleLocation != null)
                        {
                            FamilySymbol CastSymbol2 = new RectWallCast().findWall_CastSymbol(doc, Wall_Cast, pickPipe);
                            Parameter pipeWidth = m.getPipeWidth(pickPipe);
                            Parameter pipeHigh = m.getPipeHeight(pickPipe);
                            instance = doc.Create.NewFamilyInstance(HoleLocation, CastSymbol2, level, StructuralType.NonStructural);
                            //參數檢查
                            List<string> paraNameToCheck = new List<string>()
                                {
                                   "開口寬","開口高","開口長","系統別","TTOP","TCOP","TBOP","BBOP","BCOP","BTOP"
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
                            Parameter pipeSlope = pickPipe.LookupParameter("斜度");
                            if (pipeSlope != null && pipeSlope.AsDouble() != 0)
                            {
                                double pipeStartHeight = pickPipe.get_Parameter(BuiltInParameter.RBS_START_OFFSET_PARAM).AsDouble();
                                double pipeEndHeight = pickPipe.get_Parameter(BuiltInParameter.RBS_END_OFFSET_PARAM).AsDouble();
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
                            double pipeHeight = pickPipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble();
                            instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(pipeHeight - castDiameter + slopeOffset);
                            double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
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
                            int index_lowLevel = levelNames.IndexOf(level.Name);
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
                            double basicWallHeight = topLevel.Elevation - level.Elevation;
                            double BBOP_toSet = pipeHeight - outterDiameter / 2;
                            double BCOP_toSet = pipeHeight;
                            double BTOP_toSet = pipeHeight + outterDiameter / 2;
                            double TCOP_toSet = basicWallHeight - BCOP_toSet;
                            double TBOP_toSet = TCOP_toSet + outterDiameter / 2;
                            double TTOP_toSet = TCOP_toSet - outterDiameter / 2;
                            #region 早期算法，連同降板一起計算
                            //double BBOP_toSet = pipeHeight - wallBaseOffset - outterDiameter / 2;
                            //double BCOP_toSet = pipeHeight - wallBaseOffset;
                            //double BTOP_toSet = pipeHeight - wallBaseOffset + outterDiameter / 2;
                            //double TCOP_toSet = wallHeight - BCOP_toSet;
                            //double TBOP_toSet = TCOP_toSet + outterDiameter / 2;
                            //double TTOP_toSet = TCOP_toSet - outterDiameter / 2;
                            #endregion
                            BBOP.Set(BBOP_toSet);
                            BCOP.Set(BCOP_toSet);
                            BTOP.Set(BTOP_toSet);
                            TBOP.Set(TBOP_toSet);
                            TCOP.Set(TCOP_toSet);
                            TTOP.Set(TTOP_toSet);
                        }
                        trans.Commit();
                    }
                    #endregion 
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
            double levelHeight = element.LookupParameter("立面").AsDouble();
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
    }
    public class RectWallCast
    {
        #region
        //將穿牆套管的功能做成class管理
        //1.先找到套管的Family
        //2.用Family反查Symbol
        #endregion
        public Family WallCastSymbol(Document doc)
        {
            string internalNameWall = "CEC-穿牆開口-方";
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
                MessageBox.Show("尚未載入指定的穿牆套管元件!");
            }
            return Wall_CastType;
        }
        public FamilySymbol findWall_CastSymbol(Document doc, Family CastFamily, Element element)
        {
            FamilySymbol targetFamilySymbol = null; //用來找目標familySymbol
            if (CastFamily != null)
            {
                //因為方形是通用類型的關係，直接找到第一個就可以返回
                FamilySymbol tempSymbol = doc.GetElement(CastFamily.GetFamilySymbolIds().First()) as FamilySymbol;
                targetFamilySymbol = tempSymbol;
            }
            targetFamilySymbol.Activate();
            return targetFamilySymbol;
        }
    }
}
