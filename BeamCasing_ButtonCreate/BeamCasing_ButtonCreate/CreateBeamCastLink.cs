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

namespace BeamCasing_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreateBeamCastLink : IExternalCommand
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
                    UIApplication uiapp = commandData.Application;
                    Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                    UIDocument uidoc = commandData.Application.ActiveUIDocument;
                    Document doc = uidoc.Document;

                    #region 蒐集管與外參樑，關係為多管對一樑
                    //點選要放置穿樑套管的管段
                    ISelectionFilter linkPipefilter = new linkedPipeSelectionFilter(doc);
                    Reference refer = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkPipefilter, $"請選擇「連結模型」中欲放置穿樑套管的管段");
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTrans = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);

                    //點選要取交集的外參樑
                    ISelectionFilter beamFilter = new BeamsLinkedSelectedFilter(doc);
                    Reference pikBeamRefer = uidoc.Selection.PickObject(ObjectType.LinkedElement, beamFilter, $"請選擇穿過的樑");
                    RevitLinkInstance beamLinkInst = doc.GetElement(pikBeamRefer) as RevitLinkInstance;
                    Document linkBeamDoc = beamLinkInst.GetLinkDocument();
                    Element pickBeam = linkBeamDoc.GetElement(pikBeamRefer.LinkedElementId);
                    Transform linkTransform = beamLinkInst.GetTotalTransform();
                    LocationCurve beamLocate = pickBeam.Location as LocationCurve;
                    Line beamLine = beamLocate.Curve as Line;
                    XYZ beamDir = beamLine.Direction.Normalize();
                    XYZ beamNorDir = beamDir.CrossProduct(XYZ.BasisZ).Normalize().Negate();
                    XYZ holeDir = XYZ.BasisY;
                    double angle = holeDir.AngleTo(beamNorDir);
                    #endregion

                    Family RC_Cast;

                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("載入檔案測試");
                        RC_Cast = new BeamCast().BeamCastSymbol(doc);
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
                    //尋找連結模型中的元素_方法2
                    IList<FamilyInstance> CastList = new List<FamilyInstance>(); //創造一個裝每次被創造出來的familyinstance的容器，用來以bounding box計算bop&top
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
                                intersectLength += intersection.GetCurveSegment(i).Length;
                                Curve tempCurve = intersection.GetCurveSegment(i);
                                XYZ tempCenter = tempCurve.Evaluate(0.5, true);
                                double tempStart = tempCurve.GetEndPoint(0).Z;
                                double tempEnd = tempCurve.GetEndPoint(1).Z;
                                FamilySymbol CastSymbol2 = new BeamCast().findRC_CastSymbol(doc, RC_Cast, linkedPipe);
                                instance = doc.Create.NewFamilyInstance(tempCenter, CastSymbol2, topLevel, StructuralType.NonStructural);

                                List<string> paraNameToCheck = new List<string>()
                                {
                                   "開口長","系統別","TTOP","TCOP","TBOP","BBOP","BCOP","BTOP"
                                };

                                foreach (string item in paraNameToCheck)
                                {
                                    if (!checkPara(instance, item))
                                    {
                                        MessageBox.Show($"執行失敗，請檢查{instance.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                                        return Result.Failed;
                                    }
                                }

                                //調整長度與高度
                                instance.LookupParameter("開口長").Set(intersection.GetCurveSegment(i).Length + 2 / 30.48); //套管前後加兩公分
                                double floorHeight = topLevel.ProjectElevation - lowLevel.ProjectElevation;
                                double adjust = instance.LookupParameter("管外半徑").AsDouble();
                                double toMove = tempCenter.Z - topLevel.ProjectElevation + adjust;
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
                                Curve beamCurve = beamLocate.Curve;
                                XYZ beamStart = beamCurve.GetEndPoint(0);
                                XYZ beamEnd = beamCurve.GetEndPoint(1);
                                beamStart = TransformPoint(beamStart, linkTransform);
                                beamEnd = TransformPoint(beamEnd, linkTransform);
                                beamEnd = new XYZ(beamEnd.X, beamEnd.Y, beamStart.Z);
                                Line beamCrvProject = Line.CreateBound(beamStart, beamEnd);

                                Line intersectProject = Line.CreateBound(projectStart, projectEndAdj);
                                double degree = 0.0;
                                degree = -basePoint.AngleTo(beamCrvProject.Direction) - Math.PI / 2;
                                double degree2 = Math.Abs(-basePoint.AngleTo(beamCrvProject.Direction) - Math.PI / 2) * 180 / Math.PI;
                                double degreeCheck = Math.Round(degree2, 0);
                                //instance.Location.Rotate(Axis, degree);
                                instance.Location.Rotate(Axis, angle);

                                //針對已在樑中的穿樑套管做檢核
                                double casrCreatedWidth = instance.get_BoundingBox(null).Max.Z - instance.get_BoundingBox(null).Min.Z;
                                LocationPoint castCreatedLocate = instance.Location as LocationPoint;
                                XYZ castCreatedXYZ = castCreatedLocate.Point;

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
                            //return Result.Failed;
                            MessageBox.Show("管沒有和任何的樑交集，請重新調整!");
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
    }
}
