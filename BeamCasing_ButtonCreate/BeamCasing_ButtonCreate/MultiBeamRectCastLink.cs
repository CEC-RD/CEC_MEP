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

namespace BeamCasing_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class MultiBeamRectCastLink : IExternalCommand
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
                IList<Reference> linkedPickPipeRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, linkedPipeFilter, "請選擇「連結模型」中貫穿樑的管");
                List<Element> pickPipes = new List<Element>();
                foreach (Reference refer in linkedPickPipeRefs)
                {
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTransform = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                    pickPipes.Add(linkedPipe);
                }

                //拿到樑外參與Transform
                ISelectionFilter linkedBeamFilter= new BeamsLinkedSelectedFilter(doc);
                Reference pickBeamRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedBeamFilter, "請選擇「連結模型」中被貫穿的樑");
                RevitLinkInstance pickBeam = doc.GetElement(pickBeamRef) as RevitLinkInstance;
                Transform linkTransform = pickBeam.GetTotalTransform();
                Element linkedBeam = pickBeam.GetLinkDocument().GetElement(pickBeamRef.LinkedElementId);
                LocationCurve beamLocate = linkedBeam.Location as LocationCurve;
                Line beamLine = beamLocate.Curve as Line;
                XYZ beamDir = beamLine.Direction.Normalize();
                XYZ beamNorDir = beamDir.CrossProduct(XYZ.BasisZ).Normalize().Negate();
                XYZ holeDir = XYZ.BasisY;
                double angle = holeDir.AngleTo(beamNorDir);

                //取得管線的curve 和樓層
                MEPCurve pipeCrv = pickPipes.First() as MEPCurve;
                LocationCurve pipeLocate = pipeCrv.Location as LocationCurve;
                Curve pipeCurve = pipeLocate.Curve;
                Level level = pipeCrv.ReferenceLevel;

                //取得套管長度
                Solid beamSolid = singleSolidFromElement(linkedBeam);
                beamSolid = SolidUtils.CreateTransformed(beamSolid, linkTransform);
                SolidCurveIntersectionOptions option = new SolidCurveIntersectionOptions();
                SolidCurveIntersection intersect = beamSolid.IntersectWithCurve(pipeCurve, option);
                if (intersect.SegmentCount == 0)
                {
                    MessageBox.Show("管並未和樑交集，請重新選擇");
                    return Result.Failed;
                }
                if (level == null)
                {
                    MessageBox.Show("請檢查管的參考樓層");
                    return Result.Failed;
                }
                double holeLength = UnitUtils.ConvertFromInternalUnits(intersect.GetCurveSegment(0).Length, unitType) + 20;
                //抓到模型中所有的樓層元素，依照樓高排序。要找到位於他上方的樓層
                FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                levelCollector.WherePasses(level_Filter).WhereElementIsNotElementType().ToElements();

                List<string> levelNames = new List<string>();
                Level lowLevel = null; //管在的樓層為下樓層
                if (lowLevel == null)
                {
                    string LevelName = pipeCrv.ReferenceLevel.Name;
                    FilteredElementCollector levelFilter = new FilteredElementCollector(doc).OfClass(typeof(Level));
                    foreach (Level le in levelFilter)
                    {
                        if (le.Name == LevelName)
                        {
                            lowLevel = le;
                        }
                    }
                }
                else
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

                //求出各個管線和樑的交集處
                double elevation = topLevel.Elevation;
                List<XYZ> intersectHeight = new List<XYZ>();
                List<XYZ> intersectWidth = new List<XYZ>();
                foreach (Element pickPipe in pickPipes)
                {
                    double half_pipeHeight = getPipeHeight(pickPipe).AsDouble() / 2;
                    double half_pipeWidth = getPipeWidth(pickPipe).AsDouble() / 2;
                    XYZ widthOffset = half_pipeWidth * beamDir;
                    XYZ holeLocation = GetHoleLocation(linkedBeam, pickPipe, linkTransform);
                    if (holeLocation == null)
                    {
                        MessageBox.Show("執行失敗，請確認選中的管是否都有和樑交集!");
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

                //幫list中的資料重新排序
                double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                double castHeight = 0.0;
                double castWidth = 0.0;
                intersectHeight = intersectHeight.OrderBy(pt => pt.Z).ToList();
                castHeight = Math.Abs(intersectHeight.First().Z - intersectHeight.Last().Z);
                if (Math.Abs(beamDir.X) > Math.Abs(beamDir.Y))
                {
                    intersectWidth = intersectWidth.OrderBy(pt => pt.X).ToList();
                    castWidth = Math.Abs(intersectWidth.First().X - intersectWidth.Last().X);
                }
                else if (Math.Abs(beamDir.X) < Math.Abs(beamDir.Y))
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
                XYZ targetPt = new XYZ((widthPt1.X + widthPt2.X) / 2, (widthPt1.Y + widthPt2.Y) / 2, (heightPt1.Z + heightPt2.Z) / 2);
                double offsetToSet = (heightPt1.Z + heightPt2.Z) / 2 - elevation;

                Family Beam_Cast;
                FamilyInstance instance = null;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("載入檔案測試");
                    Beam_Cast = new RectBeamCast().BeamCastFamily(doc);
                    tx.Commit();
                }


                using (TransactionGroup transactionGroup = new TransactionGroup(doc))
                {
                    transactionGroup.Start("放置多管共用套管");

                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置套管");
                        FamilySymbol CastSymbol2 = new RectBeamCast().findBeam_CastSymbol(doc, Beam_Cast);
                        instance = doc.Create.NewFamilyInstance(targetPt, CastSymbol2, topLevel, StructuralType.NonStructural);
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
#if RELEASE2019
                        instance.LookupParameter("偏移").Set(offsetToSet + castDiameter);
#else
                        instance.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(offsetToSet + castDiameter);
#endif
                        trans.Commit();

                        //設定BOP、TOP
                        if (intersect.SegmentCount > 0)
                        {
                            using (Transaction trans2 = new Transaction(doc))
                            {
                                trans2.Start("調整參數");
                                Solid tempBeam = beamSolid; //如果樑有切割到，則對樑進行計算
                                XYZ instance_Max = instance.get_BoundingBox(null).Max;
                                XYZ instance_Min = instance.get_BoundingBox(null).Min;
                                //MessageBox.Show($"套管的上部座標為：{instance_Max}，下部座標為：{instance_Min}");
                                //double instanceHeight = instance_Max.Z - instance_Min.Z; //穿樑套管的高度

                                LocationPoint instLocate = instance.Location as LocationPoint;
                                XYZ instPt = instLocate.Point;
                                //MessageBox.Show(instPt.Z.ToString());

                                //針對每個實體
                                double normal_BeamHeight = UnitUtils.ConvertToInternalUnits(1500, unitType);
                                XYZ tempCenter_Up = new XYZ(instPt.X, instPt.Y, instPt.Z + normal_BeamHeight);
                                XYZ tempCenter_Dn = new XYZ(instPt.X, instPt.Y, instPt.Z - normal_BeamHeight);
                                Curve vertiaclLine = Line.CreateBound(tempCenter_Dn, tempCenter_Up);

                                SolidCurveIntersection castIntersect = beamSolid.IntersectWithCurve(vertiaclLine, option);
                                Curve castIntersect_Crv = castIntersect.GetCurveSegment(0);
                                XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                                XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);

                                //MessageBox.Show($"切到的上部Z值為：{intersect_UP}，下部Z值為：{intersect_DN}");
                                double castCenter_Z = (instance_Max.Z + instance_Min.Z) / 2;
                                double TTOP = intersect_UP.Z - instance_Max.Z;
                                double BTOP = instance_Max.Z - intersect_DN.Z;
                                double TCOP = intersect_UP.Z - castCenter_Z;
                                double BCOP = castCenter_Z - intersect_DN.Z;
                                double TBOP = intersect_UP.Z - instance_Min.Z;
                                double BBOP = instance_Min.Z - intersect_DN.Z;
                                //MessageBox.Show(instance_Max.Z.ToString());

                                instance.LookupParameter("TTOP").Set(TTOP);
                                instance.LookupParameter("BTOP").Set(BTOP);
                                instance.LookupParameter("TCOP").Set(TCOP);
                                instance.LookupParameter("BCOP").Set(BCOP);
                                instance.LookupParameter("TBOP").Set(TBOP);
                                instance.LookupParameter("BBOP").Set(BBOP);
                                trans2.Commit();
                            }
                        }
                    }
                    transactionGroup.Assimilate();
                }
            }
            catch
            {
                //MessageBox.Show("執行失敗");
                return Result.Failed;
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
