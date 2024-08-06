#region using namespace
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
#endregion

namespace BeamCasing_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreateRectBeamCastLink : IExternalCommand
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


                    //點選要放置穿樑套管的管段
                    ISelectionFilter linkPipefilter = new linkedPipeSelectionFilter(doc);
                    Reference refer = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkPipefilter, $"請選擇「連結模型」中欲放置穿樑套管的管段");
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTrans = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                    //點選要取交集的外參樑
                    ISelectionFilter beamFilter = new BeamsLinkedSelectedFilterST(doc);
                    Reference pikBeamRefer = uidoc.Selection.PickObject(ObjectType.LinkedElement, beamFilter, $"請選擇穿過的樑");
                    RevitLinkInstance beamLinkInst = doc.GetElement(pikBeamRefer) as RevitLinkInstance;
                    Document linkBeamDoc = beamLinkInst.GetLinkDocument();
                    Element pickBeam = linkBeamDoc.GetElement(pikBeamRefer.LinkedElementId);
                    Transform linkTransform = beamLinkInst.GetTotalTransform();

                    //load Family target
                    Family rectCast;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        rectCast = new RectBeamCast().BeamCastFamily(doc);
                        tx.Commit();
                    }


                    //抓到模型中所有的樓層元素，依照樓高排序。要找到位於他上方的樓層
                    FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                    ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                    levelCollector.WherePasses(level_Filter).WhereElementIsNotElementType().ToElements();
                    List<string> levelNames = new List<string>(); //用名字來確認篩選排序
                    MEPCurve pipeCrv = linkedPipe as MEPCurve;
                    //Level lowLevel = pipeCrv.ReferenceLevel; //管在的樓層為下樓層
                    Level lowLevel = null; //管在的樓層為下樓層


                    string LevelName = pipeCrv.ReferenceLevel.Name;
                    Level sourceLevel = pipeCrv.ReferenceLevel;
                    FilteredElementCollector levelFilter = new FilteredElementCollector(doc).OfClass(typeof(Level));
                    foreach (Level le in levelFilter)
                    {
                        if (le.Name == LevelName || le.ProjectElevation == sourceLevel.ProjectElevation)
                        {
                            lowLevel = le;
                        }
                    }

                    if (lowLevel == null)
                    {
                        MessageBox.Show("請確認「連結模型」中的樓層命名原則是否和本機端一致");
                        return Result.Failed;
                    }
                    List<Element> level_List = levelCollector.OrderBy(x => sortLevelbyHeight(x)).ToList();

                    for (int i = 0; i < level_List.Count(); i++)
                    {
                        Level le = level_List[i] as Level;
                        levelNames.Add(le.Name);
                    }
                    int index_lowLevel = levelNames.IndexOf(lowLevel.Name);
                    int index_topLevel = index_lowLevel + 1;
                    Level topLevel = null;
                    if (index_topLevel < level_List.Count())
                    {
                        topLevel = level_List[index_topLevel] as Level;
                    }
                    else if (topLevel == null)
                    {
                        message = "管的上方沒有樓層，無法計算穿樑套管偏移值";
                        return Result.Failed;
                    }

                    FamilyInstance instance = null;
                    int intersectCount = 0;
                    double intersectLength = 0;
                    int totalIntersectCount = 0;
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置穿樑套管");
                        Solid solid = singleSolidFromElement(pickBeam);
                        solid = SolidUtils.CreateTransformed(solid, linkTransform);
                        if (null != solid)
                        {
                            SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();
                            LocationCurve locationCurve = linkedPipe.Location as LocationCurve;
                            Curve pipeCurve = locationCurve.Curve;
                            SolidCurveIntersection intersection = solid.IntersectWithCurve(pipeCurve, options);
                            intersectCount = intersection.SegmentCount;
                            totalIntersectCount += intersectCount;
                            for (int i = 0; i < intersectCount; i++)
                            {
                                Curve tempCurve = intersection.GetCurveSegment(i);
                                intersectLength += tempCurve.Length;
                                XYZ tempCenter = tempCurve.Evaluate(0.5, true);
                                double tempStart = tempCurve.GetEndPoint(0).Z;
                                double tempEnd = tempCurve.GetEndPoint(1).Z;
                                FamilySymbol castSymbol = new RectBeamCast().findBeam_CastSymbol(doc, rectCast);

                                instance = doc.Create.NewFamilyInstance(tempCenter, castSymbol, topLevel, StructuralType.NonStructural);
                                List<string> paraNameToCheck = new List<string>()
                                {
                                  "開口寬","開口高","開口長","系統別","TTOP","TCOP","TBOP","BBOP","BCOP","BTOP"
                                };

                                foreach (string item in paraNameToCheck)
                                {
                                    if (!checkPara(instance, item))
                                    {
                                        MessageBox.Show($"執行失敗，請檢查{instance.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
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
                                    double distToStart = tempCenter.DistanceTo(startPt);
                                    if (pipeStartHeight >= pipeEndHeight)
                                    {
                                        slopeOffset = -(distToStart * pipeSlope.AsDouble());
                                    }
                                    else if (pipeStartHeight < pipeEndHeight)
                                    {
                                        slopeOffset = distToStart * pipeSlope.AsDouble();
                                    }
                                }

                                //設定方形開口的參數
                                double castOffset = UnitUtils.ConvertToInternalUnits(100, unitType);
                                Parameter pipeWidth = getPipeWidth(linkedPipe);
                                Parameter pipeHigh = getPipeHeight(linkedPipe);
                                double widthToSet = pipeWidth.AsDouble() + castOffset;
                                double heightToSet = pipeHigh.AsDouble() + castOffset;
                                instance.LookupParameter("開口高").Set(heightToSet);
                                instance.LookupParameter("開口寬").Set(widthToSet);
                                double beamWidth = getBeamWidth(pickBeam);
                                //2022.08.05新增 更改取得樑參數的方法
                                //instance.LookupParameter("開口長").Set(intersection.GetCurveSegment(i).Length + 2 / 30.48); //套管前後加兩公分
                                instance.LookupParameter("開口長").Set(beamWidth + 2 / 30.48);
                                double floorHeight = topLevel.Elevation - lowLevel.Elevation;
                                double toMove = tempCenter.Z - topLevel.Elevation + heightToSet / 2;
                                //double toMove = pickPipe.LookupParameter("偏移").AsDouble() - floorHeight + adjust;
#if RELEASE2019
                                instance.LookupParameter("偏移").Set(toMove);
#else
                                instance.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(toMove);
#endif
                                //旋轉角度
                                XYZ axisPt1 = new XYZ(tempCenter.X, tempCenter.Y, tempCenter.Z);
                                XYZ axisPt2 = new XYZ(tempCenter.X, tempCenter.Y, tempCenter.Z + 1);
                                XYZ basePoint = new XYZ(0, tempCenter.Y, 0);
                                Line Axis = Line.CreateBound(axisPt1, axisPt2);
                                XYZ projectStart = intersection.GetCurveSegment(i).GetEndPoint(0);
                                XYZ projectEnd = intersection.GetCurveSegment(i).GetEndPoint(1);
                                XYZ projectEndAdj = new XYZ(projectEnd.X, projectEnd.Y, projectStart.Z);


                                //重新以樑的locationCurve 計算旋轉角度
                                LocationCurve beamLocate = pickBeam.Location as LocationCurve;
                                Curve beamCurve = beamLocate.Curve;
                                XYZ beamStart = beamCurve.GetEndPoint(0);
                                XYZ beamEnd = beamCurve.GetEndPoint(1);
                                beamStart = TransformPoint(beamStart, linkTransform);
                                beamEnd = TransformPoint(beamEnd, linkTransform);
                                beamEnd = new XYZ(beamEnd.X, beamEnd.Y, beamStart.Z);
                                Line beamCrvProject = Line.CreateBound(beamStart, beamEnd);
                                Line intersectProject = Line.CreateBound(projectStart, projectEndAdj);
                                double degree = 0.0;
                                //degree = basePoint.AngleTo(intersectProject.Direction);
                                degree = -basePoint.AngleTo(beamCrvProject.Direction) - Math.PI / 2;
                                double degree2 = Math.Abs(-basePoint.AngleTo(beamCrvProject.Direction) - Math.PI / 2) * 180 / Math.PI;
                                double degreeCheck = Math.Round(degree2, 0);
                                instance.Location.Rotate(Axis, degree);


                                //設定BOP、TOP
                                if (intersectCount > 0)
                                {
                                    Solid tempBeam = solid; //如果樑有切割到，則對樑進行計算
                                    XYZ tempBeam_Max = tempBeam.GetBoundingBox().Max;
                                    XYZ tempBeam_Min = tempBeam.GetBoundingBox().Min;

                                    XYZ instance_Max = instance.get_BoundingBox(null).Max;
                                    XYZ instance_Min = instance.get_BoundingBox(null).Min;
                                    double instanceHeight = instance_Max.Z - instance_Min.Z; //穿樑套管的高度

                                    //針對每個實體
                                    XYZ tempCenter_Up = new XYZ(tempCenter.X, tempCenter.Y, tempCenter.Z + 50);
                                    XYZ tempCenter_Dn = new XYZ(tempCenter.X, tempCenter.Y, tempCenter.Z - 50);
                                    Curve vertiaclLine = Line.CreateBound(tempCenter_Dn, tempCenter_Up);

                                    SolidCurveIntersection castIntersect = solid.IntersectWithCurve(vertiaclLine, options);
                                    Curve castIntersect_Crv = castIntersect.GetCurveSegment(0);
                                    XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                                    XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);

                                    double castCenter_Z = (instance_Max.Z + instance_Min.Z) / 2;
                                    double TTOP = intersect_UP.Z - instance_Max.Z;
                                    double BTOP = instance_Max.Z - intersect_DN.Z;
                                    double TCOP = intersect_UP.Z - castCenter_Z;
                                    double BCOP = castCenter_Z - intersect_DN.Z;
                                    double TBOP = intersect_UP.Z - instance_Min.Z;
                                    double BBOP = instance_Min.Z - intersect_DN.Z; ;

                                    instance.LookupParameter("TTOP").Set(TTOP);
                                    instance.LookupParameter("BTOP").Set(BTOP);
                                    instance.LookupParameter("TCOP").Set(TCOP);
                                    instance.LookupParameter("BCOP").Set(BCOP);
                                    instance.LookupParameter("TBOP").Set(TBOP);
                                    instance.LookupParameter("BBOP").Set(BBOP);


                                    //設定1/4的穿樑原則警告
                                    double beamHeight = castIntersect_Crv.Length; //樑在那個斷面的高度
                                    double alertValue = beamHeight / 5;
                                    double min_alertValue = UnitUtils.ConvertToInternalUnits(200, unitType);//除了1/4或1/3的保護層限制之外，也有最小保護層的限制
                                    if (alertValue < min_alertValue)
                                    {
                                        alertValue = min_alertValue;
                                    }
                                }

                            }
                        }
                        if (totalIntersectCount == 0)
                        {
                            //message = "管沒有和任何的樑交集，請重新調整!";
                            //elements.Insert(pickPipe);
                            MessageBox.Show("管沒有和任何的樑交集，請重新調整!");
                            return Result.Failed;
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
        public double sortLevelbyHeight(Element element)
        {
            Level tempLevel = element as Level;
            double levelHeight = tempLevel.Elevation;
            return levelHeight;
        }
        public IList<Solid> GetTargetSolids(Element element)
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
        public Solid singleSolidFromElement(Element inputElement)
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
        public double getBeamWidth(Element beam)
        {
            FamilyInstance beamInst = beam as FamilyInstance;
            FamilySymbol beamSymbol = beamInst.Symbol;
            List<string> paraNameLst = new List<string>() { "樑寬度", "梁寬度", "樑寬", "梁寬", "寬度", "寬", "B", "W", "b" };
            List<double> paraValue = new List<double>();
            double targetValue = 0.0;
            foreach (string st in paraNameLst)
            {
                Parameter tempPara = beamSymbol.LookupParameter(st);
                if (tempPara != null && tempPara.AsDouble() != 0)
                {
                    paraValue.Add(tempPara.AsDouble());
                }
            }
            targetValue = paraValue.Max();
            if (targetValue == 0)
            {
                MessageBox.Show("請檢察樑中的「寬度」參數是否有誤，無法更新套管長度(或套管長不可為0)!");
            }
            return targetValue;
        }
    }
}
