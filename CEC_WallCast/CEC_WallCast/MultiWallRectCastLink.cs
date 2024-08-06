#region Namespaces
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
#endregion

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class MultiWallRectCastLink : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;
                ISelectionFilter linkedPipeFilter = new linkedPipeSelectionFilter(doc);
                IList<Reference> linkedPickPipeRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, linkedPipeFilter, "請選擇「連結模型」中貫穿牆的管");
                List<Element> pickPipes = new List<Element>();
                foreach (Reference refer in linkedPickPipeRefs)
                {
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTransform = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                    pickPipes.Add(linkedPipe);
                    //output += $"{linkedPipe.Name}\n";
                }
                //拿到整份牆外參檔&Transform
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
                XYZ wallDir = new XYZ();
                if (wallLine != null)
                {
                    isLiner = true;
                    wallDir = wallLine.Direction.Normalize();
                    XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate(); //這段要再看一下
                    angle = holeDir.AngleTo(wallNorDir);
                }
                else
                {
                    //重新塑造牆的向量，但這做法有點問題
                    Line tempLine = Line.CreateBound(wallCrv.GetEndPoint(0), wallCrv.GetEndPoint(1));
                    wallDir = tempLine.Direction.Normalize();
                }
                MEPCurve linkPipeCrv = pickPipes.First() as MEPCurve;
                LocationCurve pipeLocate = linkPipeCrv.Location as LocationCurve;
                Curve pipeCurve = pipeLocate.Curve;
                //取得管線的參考樓層，要因應外參管和本地端管因為樓層元素不一樣的狀況進行判斷
                Level pipeLevel = null;

                string LevelName = linkPipeCrv.ReferenceLevel.Name;
                Level sourceLevel = linkPipeCrv.ReferenceLevel;
                FilteredElementCollector levelFilter = new FilteredElementCollector(doc).OfClass(typeof(Level));
                foreach (Level le in levelFilter)
                {
                    if (le.Name == LevelName || le.ProjectElevation ==sourceLevel.ProjectElevation)
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
                List<XYZ> intersectHeight = new List<XYZ>();
                List<XYZ> intersectWidth = new List<XYZ>();
                foreach (Element pickPipe in pickPipes)
                {
                    double half_pipeHeight = getPipeHeight(pickPipe).AsDouble() / 2;
                    double half_pipeWidth = getPipeWidth(pickPipe).AsDouble() / 2;
                    XYZ widthOffset = half_pipeWidth * wallDir;
                    XYZ holeLocation = GetHoleLocation(linkedWall, pickPipe, linkTransform);
                    if (holeLocation == null)
                    {
                        MessageBox.Show("執行失敗，請確認選中的管是否都有和牆交集!");
                        return Result.Failed;
                    }
                    XYZ temp_intersectUp = new XYZ(holeLocation.X, holeLocation.Y, holeLocation.Z + half_pipeHeight);
                    XYZ temp_intersectDn = new XYZ(holeLocation.X, holeLocation.Y, holeLocation.Z - half_pipeHeight);
                    XYZ temp_intersectRight = new XYZ(holeLocation.X + widthOffset.X, holeLocation.Y + widthOffset.Y, holeLocation.Z);
                    XYZ temp_intersectLeft = new XYZ(holeLocation.X - widthOffset.X, holeLocation.Y - widthOffset.Y, holeLocation.Z);
                    intersectHeight.Add(temp_intersectUp);
                    intersectHeight.Add(temp_intersectDn);
                    intersectWidth.Add(temp_intersectRight);
                    intersectWidth.Add(temp_intersectLeft);
                }
                double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                double castHeight = 0.0;
                double castWidth = 0.0;
                intersectHeight = intersectHeight.OrderBy(pt => pt.Z).ToList();
                castHeight = Math.Abs(intersectHeight.First().Z - intersectHeight.Last().Z);
                if (Math.Abs(wallDir.X) > Math.Abs(wallDir.Y))
                {
                    intersectWidth = intersectWidth.OrderBy(pt => pt.X).ToList();
                    castWidth = Math.Abs(intersectWidth.First().X - intersectWidth.Last().X);
                }
                else if (Math.Abs(wallDir.X) < Math.Abs(wallDir.Y))
                {
                    intersectWidth = intersectWidth.OrderBy(pt => pt.Y).ToList();
                    castWidth = Math.Abs(intersectWidth.First().Y - intersectWidth.Last().Y);
                }
                double castOffset = UnitUtils.ConvertToInternalUnits(100, unitType);
                castHeight = castHeight + castOffset;
                castWidth = castWidth + castOffset;

                //利用這點，找到要放置方套管的點位
                XYZ widthPt1 = intersectWidth.First();
                XYZ widthPt2 = intersectWidth.Last();
                XYZ heightPt1 = intersectHeight.First();
                XYZ heightPt2 = intersectHeight.Last();
                XYZ targetPt = new XYZ((widthPt1.X + widthPt2.X) / 2, (widthPt1.Y + widthPt2.Y) / 2, (heightPt1.Z + heightPt2.Z) / 2 - elevation);
                //如果牆不為曲線，須單獨設定，計算曲線導數作為旋轉依據
                if (isLiner != true)
                {
                    IntersectionResult intersect = wallCrv.Project(targetPt);
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
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("放置多管共用套管");
                    FamilySymbol CastSymbol2 = new RectWallCast().findWall_CastSymbol(doc, Wall_Cast, pickPipes.First());
                    instance = doc.Create.NewFamilyInstance(targetPt, CastSymbol2, pipeLevel, StructuralType.NonStructural);
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

                    instance.LookupParameter("開口高").Set(castHeight);
                    instance.LookupParameter("開口寬").Set(castWidth);
                    instance.LookupParameter("開口長").Set(castLength);
                    //旋轉開口
                    LocationPoint loPoint = instance.Location as LocationPoint;
                    XYZ newPoint = new XYZ(targetPt.X, targetPt.Y, targetPt.Z + 10);
                    Line axis = Line.CreateBound(targetPt, newPoint);
                    loPoint.Rotate(axis, angle);

                    //調整最終高度
                    double castDiameter = castHeight / 2;
                    //double originHeight = instance.LookupParameter("偏移").AsDouble();
                    double originHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                    Parameter instZ = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
                    instZ.Set(originHeight - castDiameter);
                    //instance.LookupParameter("偏移").Set(originHeight - castDiameter);

                    //尚缺設定BBOP等參數的功能
                    Parameter BBOP = instance.LookupParameter("BBOP");
                    Parameter BCOP = instance.LookupParameter("BCOP");
                    Parameter BTOP = instance.LookupParameter("BTOP");
                    Parameter TTOP = instance.LookupParameter("TTOP");
                    Parameter TCOP = instance.LookupParameter("TCOP");
                    Parameter TBOP = instance.LookupParameter("TBOP");
                    double outterDiameter = instance.LookupParameter("開口高").AsDouble();
                    double wallBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
                    double wallHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                    double instHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                    //double instHeight = instance.LookupParameter("偏移").AsDouble();


                    FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                    ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                    levelCollector.WherePasses(level_Filter).WhereElementIsNotElementType().ToElements();
                    List<string> levelNames = new List<string>();
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
                    //預計寫入的計算值
                    double basicWallHeight = topLevel.Elevation - pipeLevel.Elevation;
                    double BBOP_toSet = instHeight;
                    double BCOP_toSet = instHeight + outterDiameter / 2;
                    double BTOP_toSet = instHeight + outterDiameter;
                    double TCOP_toSet = basicWallHeight - BCOP_toSet;
                    double TBOP_toSet = TCOP_toSet + outterDiameter / 2;
                    double TTOP_toSet = TCOP_toSet - outterDiameter / 2;
                    BBOP.Set(BBOP_toSet);
                    BCOP.Set(BCOP_toSet);
                    BTOP.Set(BTOP_toSet);
                    TBOP.Set(TBOP_toSet);
                    TCOP.Set(TCOP_toSet);
                    TTOP.Set(TTOP_toSet);

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
            double levelHeight = tempLevel.Elevation;
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
