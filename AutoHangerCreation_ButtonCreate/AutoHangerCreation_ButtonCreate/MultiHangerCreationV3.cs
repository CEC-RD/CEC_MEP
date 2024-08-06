using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Autodesk.Revit.DB.Structure;
using System;
using Autodesk.Revit.UI.Selection;
using System.Linq;
using BuildingCoder;

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class MultiHangerCreationV3 : IExternalCommand
    {
        //放置多管吊架
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Centimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Counter.count += 1;
            //蒐集所有的管->包含外參管
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
            ISelectionFilter linkedPipeFilter = new linkedPipeSelectionFilter(doc);
            IList<Reference> pickElements_Refer = uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, $"請選擇「本地端」的管");
            //先蒐集實體管
            //List<Element> pickElements = new List<Element>();
            List<ElementWithTrans> pickElements = new List<ElementWithTrans>();
            foreach (Reference reference in pickElements_Refer)
            {
                Element element = doc.GetElement(reference.ElementId);
                LocationCurve pipeLocate = element.Location as LocationCurve;
                Curve pipeCrv = pipeLocate.Curve as Line;
                ElementWithTrans localEle = new ElementWithTrans { pipe = element, locateCrv = pipeCrv };
                pickElements.Add(localEle);
                //pickElements.Add(element);
            }
            IList<Reference> linkedPickPipeRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, linkedPipeFilter, "請選擇「連結模型」的管(選填)");
            foreach (Reference refer in linkedPickPipeRefs)
            {
                RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                Transform pipeLinkedTransform = pipeLinkedInst.GetTotalTransform();
                Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                LocationCurve linkedLocate = linkedPipe.Location as LocationCurve;
                Curve MEPCrv = linkedLocate.Curve;
                XYZ startPoint = MEPCrv.GetEndPoint(0);
                XYZ endPoint = MEPCrv.GetEndPoint(1);
                startPoint =method.TransformPoint(startPoint, pipeLinkedTransform);
                endPoint = method.TransformPoint(endPoint, pipeLinkedTransform);
                Line tempCrv = Line.CreateBound(startPoint, endPoint);
                ElementWithTrans linkedEle = new ElementWithTrans { pipe = linkedPipe, locateCrv = tempCrv };
                pickElements.Add(linkedEle);
            }
            while (true)
            {
                try
                {
                    Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, linkedPipeFilter, "請點選欲放置管架的位置");
                    XYZ firstPoint = refer.GlobalPoint;
                    //XYZ firstPoint = uidoc.Selection.PickPoint();
                    pickElements = pickElements.OrderBy(x => getPipeIntersection(x.pipe, x.locateCrv, firstPoint)).ToList();

                    //先蒐集所有的樓層後，按高度排序
                    FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
                    ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                    levelCollector.WherePasses(level_Filter).WhereElementIsNotElementType().ToElements();
                    List<Element> level_List = levelCollector.OrderBy(x => sortLevelbyHeight(x)).ToList();

                    List<string> levelNames = new List<string>();
                    Level lowLevel = null;
                    if (lowLevel == null)
                    {
                        MEPCurve firstMEP = pickElements.First().pipe as MEPCurve;
                        string LevelName = firstMEP.ReferenceLevel.Name;
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


                    Family targetFamily = findHangerFamily(doc);
                    //FamilySymbol targetSymbol = doc.GetElement(targetFamily.GetFamilySymbolIds().First()) as FamilySymbol;

                    using (TransactionGroup transactionGroup = new TransactionGroup(doc))
                    {
                        transactionGroup.Start("放置多管吊架");
                        using (Transaction trans = new Transaction(doc))
                        {
                            trans.Start("放置多管吊架");
                            Family multiHanger = findHangerFamily(doc);
                            FamilySymbol symbol = doc.GetElement(multiHanger.GetFamilySymbolIds().First()) as FamilySymbol;
                            string APIcheck = "API識別名稱";
                            double adjustValue = 0.0;
                            //symbol.Activate();
                            //以集合中的第一條線作為放置的基準
                            Element firstEle = pickElements.First().pipe;
                            double outterDia = getOutterDia(firstEle);
                            if (symbol.LookupParameter(APIcheck).AsString().Contains("管底"))
                            {
                                adjustValue = outterDia / 2;
                            }
                            else if (symbol.LookupParameter(APIcheck).AsString().Contains("管頂"))
                            {
                                adjustValue = -(outterDia / 2);
                            }
                            XYZ targetPt = pickElements.First().locateCrv.Project(firstPoint).XYZPoint;
                            firstPoint = new XYZ(firstPoint.X, firstPoint.Y, targetPt.Z);

                            XYZ newTargetPt = new XYZ(targetPt.X, targetPt.Y, targetPt.Z - lowLevel.ProjectElevation - adjustValue);
                            FamilyInstance instance = doc.Create.NewFamilyInstance(newTargetPt, symbol, lowLevel, StructuralType.NonStructural);

                            //針對高程必須在設定一次，否則有時吊架高程會跑掉
                            Parameter elevationPara = instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                            if (elevationPara == null) MessageBox.Show("元件高程值設定失敗");
                            elevationPara.Set(targetPt.Z - lowLevel.ProjectElevation - adjustValue);

                            //設定標稱管徑與間距
                            int pipeCount = pickElements.Count();
                            List<string> defaultDia = new List<string> { "標稱直徑1", "標稱直徑2", "標稱直徑3", "標稱直徑4", "標稱直徑5", "標稱直徑6", "標稱直徑7", "標稱直徑8", "標稱直徑9", "標稱直徑10" };
                            List<string> defaultStep = new List<string> { "管間距1", "管間距2", "管間距3", "管間距4", "管間距5", "管間距6", "管間距7", "管間距8", "管間距9" };
                            //List<string> diaPara = new List<string> { "標稱直徑1", "標稱直徑2", "標稱直徑3", "標稱直徑4", "標稱直徑5", "標稱直徑6", "標稱直徑7", "標稱直徑8", "標稱直徑9", "標稱直徑10" };
                            //List<string> stepPara = new List<string> { "管間距1", "管間距2", "管間距3", "管間距4", "管間距5", "管間距6", "管間距7", "管間距8", "管間距9" };
                            List<string> diaPara = new List<string> { };
                            List<string> stepPara = new List<string> { };

                            #region 以迴圈整理要寫入的參數欄位
                            foreach (string st in defaultDia)
                            {
                                if (method.checkPara(instance, st)) diaPara.Add(st);
                            }
                            foreach(string st in defaultStep)
                            {
                                if (method.checkPara(instance, st)) stepPara.Add(st);
                            }
                            #endregion
                            List<double> diaToSet = new List<double>();
                            List<double> stepToSet = new List<double>();
                            List<Element> pickPipes = new List<Element>();
                            double toMinus = 0;
                            foreach (ElementWithTrans eTrans in pickElements)
                            {
                                double nominalDia = getPipeDia(eTrans.pipe);
                                pickPipes.Add(eTrans.pipe);
                                diaToSet.Add(nominalDia);
                                double distance = eTrans.locateCrv.Distance(firstPoint);
                                double convertUnit = UnitUtils.ConvertFromInternalUnits(distance, unitType);
                                if (toMinus != 0)
                                {
                                    double step = distance - toMinus;
                                    stepToSet.Add(step);
                                }
                                toMinus = distance;
                            }
                            //直徑設定
                            for (int i = 0; i < diaPara.Count(); i++)
                            {
                                if (i < diaToSet.Count())
                                {
                                    instance.LookupParameter(diaPara[i]).Set(diaToSet[i]);
                                }
                                else
                                {
                                    instance.LookupParameter(diaPara[i]).Set(0);
                                }
                            }
                            //間距設定
                            for (int i = 0; i < stepPara.Count(); i++)
                            {
                                if (i < stepToSet.Count())
                                {
                                    instance.LookupParameter(stepPara[i]).Set(stepToSet[i]);
                                }
                                else
                                {
                                    instance.LookupParameter(stepPara[i]).Set(0);
                                }
                            }
                            //旋轉設定
                            double angle = 0.0;
                            XYZ newPoint = new XYZ(newTargetPt.X, newTargetPt.Y, newTargetPt.Z - 10);
                            Line axis = Line.CreateBound(newTargetPt, newPoint);
                            XYZ hangerDir = XYZ.BasisY;
                            //XYZ hangerDir = new XYZ(newTargetPt.X, 0, 0);
                            Line pipeCrv = pickElements.First().locateCrv as Line;
                            XYZ pipeDir = pipeCrv.Direction.Normalize();
                            XYZ pipeNorDir = pipeDir.CrossProduct(XYZ.BasisZ).Normalize(); //這段要再看一下
                            angle = hangerDir.AngleTo(pipeNorDir);
                            LocationPoint loPoint = instance.Location as LocationPoint;
                            double half_PI = Math.PI / 2;
                            loPoint.Rotate(axis, angle - half_PI);


                            //製作一個將吊桿延伸到底版的方法，採用快速過濾器(品類)、樓層+慢速過濾器(ElementIntersect)
                            //1.繪製一條線，以此線作為篩選的機制
                            //2.利用線的boundingbox filter搭配CategortFilter
                            //3.篩選找出來的樓板
                            //4.利用CurveIntersectionSolid來計算長度
                            double adjust = UnitUtils.ConvertToInternalUnits(1000, unitType);
                            XYZ upperPoint = new XYZ(newTargetPt.X, newTargetPt.Y, newTargetPt.Z + adjust);
                            if (method.checkPara(instance, "PipeBottomToFloor"))
                            {
                                instance.LookupParameter("PipeBottomToFloor").Set(adjust);
                            }
                            else
                            {
                                MessageBox.Show("請確認元件中是否有「PipeBottomToFloor」參數");
                            }
                            method.updateHangerContent(instance, pickPipes);
                            trans.Commit();

                            ////更改策略，改以放置吊架時盡量讓他長大，利用boundingBox 去和revitLinkInstance做干涉，找到和此吊架有干涉的版後再排序
                            ////1.放置吊架時故意讓他長大
                            ////2.篩選具有版元件的revitlinkInstance
                            ////3.利用版元件和boundingbox去做干涉，後針對這些版位置的高低進行排序
                            ////4.創造線和版的實體進行交接干涉
                            trans.Start("吊架與版資訊蒐集");
                            Dictionary<ElementId, List<Element>> hangerSlabDict = method.getHangerSlabDict(doc, instance);
                            Line intersectCrv = Line.CreateBound(newTargetPt, upperPoint);
                            List<XYZ> intersectPts = new List<XYZ>();
                            if (hangerSlabDict[instance.Id].Count() == 0) return Result.Failed;
                            foreach (Element linkSlab in hangerSlabDict[instance.Id])
                            {
                                RevitLinkInstance targetLink =method.getTargetLinkedInstance(doc, linkSlab.Document.Title);
                                Transform linkedInstTrans = targetLink.GetTotalTransform();
                                Solid slabSolid = singleSolidFromElement(linkSlab);
                                slabSolid = SolidUtils.CreateTransformed(slabSolid, linkedInstTrans);
                                SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();
                                if (null != slabSolid)
                                {
                                    SolidCurveIntersection intersection = slabSolid.IntersectWithCurve(intersectCrv, options);
                                    int intersectCount = intersection.SegmentCount;
                                    if (intersectCount > 0)
                                    {
                                        Curve castIntersect_Crv = intersection.GetCurveSegment(0);
                                        XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                                        XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);
                                        intersectPts.Add(intersect_DN);
                                    }
                                }
                            }
                            intersectPts = intersectPts.OrderBy(x => x.Z).ToList();
                            double finalLength = intersectPts.First().Z - newTargetPt.Z;
                            string threadLenName = "PipeBottomToFloor";
                            instance.LookupParameter(threadLenName).Set(finalLength);
                            trans.Commit();
                        }
                        transactionGroup.Assimilate();
                    }
                }
                catch
                {
                    break;
                }
            }
            Counter.count += 1;
            return Result.Succeeded;
        }

        //public RevitLinkInstance getTargetLinkedInstance(Document doc, string linkTilte)
        //{
        //    RevitLinkInstance targetLinkInstance = null;
        //    try
        //    {
        //        ElementCategoryFilter linkedFileFilter = new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks);
        //        FilteredElementCollector linkedFileCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType();
        //        if (linkedFileCollector.Count() > 0)
        //        {
        //            foreach (RevitLinkInstance linkedInst in linkedFileCollector)
        //            {
        //                //if (linkedInst.GetLinkDocument().Title == linkTilte)
        //                if (linkedInst.Name.Contains(linkTilte))
        //                {
        //                    targetLinkInstance = linkedInst;
        //                    break;
        //                }
        //            }
        //        }
        //        else if (targetLinkInstance == null)
        //        {
        //            MessageBox.Show("未找到對應的實做Revit外參檔!!");
        //        }
        //    }
        //    catch
        //    {
        //        MessageBox.Show("找尋連結實體發生問題!");
        //    }
        //    return targetLinkInstance;
        //}
        public double getPipeIntersection(Element e, Curve curve, XYZ userSelection)
        {
            //取得管得位置資訊後進行變換
            //MEPCurve mep = e as MEPCurve;
            //LocationCurve locateCrv = mep.Location as LocationCurve;
            //Curve MEPCrv = locateCrv.Curve;
            //XYZ startPoint = MEPCrv.GetEndPoint(0);
            //XYZ endPoint = MEPCrv.GetEndPoint(1);
            //startPoint = TransformPoint(startPoint, trans);
            //endPoint = TransformPoint(endPoint, trans);
            //endPoint = new XYZ(endPoint.X, endPoint.Y, startPoint.Z);
            //Line tempCrv = Line.CreateBound(startPoint, endPoint);
            //XYZ projection = tempCrv.Project(userSelection).XYZPoint;
            //double distance = tempCrv.Distance(userSelection);
            double distance = curve.Distance(userSelection);
            return distance;
        }
      
        public Family findHangerFamily(Document doc)
        {
            string targetName = PIpeHangerSetting.Default.MultiHangerSelected;
            //string targetName = "M_多管吊架(L角鐵)";
            //string targetName = "M_多管吊架(C型鋼)";
            Family targetFamily = null;
            FilteredElementCollector elemColl = new FilteredElementCollector(doc).OfClass(typeof(Family));
            foreach (Element e in elemColl)
            {
                Family tempFam = e as Family;
                if (tempFam.Name == targetName)
                {
                    targetFamily = tempFam;
                }
            }
            if (targetFamily == null) MessageBox.Show("尚未匯入指定的管吊架元件");
            return targetFamily;
        }
        //public Dictionary<ElementId, List<Element>> getHangerSlabDict(Document doc, FamilyInstance inst)
        //{
        //    //其實應該再試試另一個方法，是以連結檔中樓板為主進行變換的方法
        //    Dictionary<ElementId, List<Element>> hangerSlabDict = new Dictionary<ElementId, List<Element>>();
        //    try
        //    {
        //        List<RevitLinkInstance> slabLinkInstances = method.findLinkInstWithSlab(doc);
        //        Transform totalTransform = null;
        //        Transform inverseTransform = null;
        //        if (slabLinkInstances.Count != 0)
        //        {
        //            foreach (RevitLinkInstance linkedInst in slabLinkInstances)
        //            {
        //                if (!hangerSlabDict.Keys.Contains(inst.Id))
        //                {
        //                    totalTransform = linkedInst.GetTotalTransform();
        //                    inverseTransform = totalTransform.Inverse;
        //                    FilteredElementCollector collectorSB = method.getAllLinkedSlab(linkedInst.GetLinkDocument());

        //                    ////針對吊架進行座標變換
        //                    Transform newTrans = Transform.Identity;
        //                    BoundingBoxXYZ instBounding = inst.get_BoundingBox(null);
        //                    Outline outLine = new Outline(inverseTransform.OfPoint(instBounding.Min), inverseTransform.OfPoint(instBounding.Max));
        //                    //快慢速過濾器結合->和樓板進行碰撞
        //                    BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
        //                    collectorSB.WherePasses(boundingBoxIntersectsFilter);
        //                    List<Element> tempList = collectorSB.ToList();
        //                    if (tempList.Count > 0)
        //                    {
        //                        hangerSlabDict.Add(inst.Id, tempList);
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("請確認吊架上方是否有樓板可供支撐!");
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        MessageBox.Show("無法判斷吊架與樓板的關係!");
        //    }
        //    return hangerSlabDict;
        //}
        public Dictionary<ElementId, List<Element>> getHangerSlabDict2(Document doc, FamilyInstance inst)
        {
            //其實應該再試試另一個方法，是以連結檔中樓板為主進行變換的方法
            Dictionary<ElementId, List<Element>> hangerSlabDict = new Dictionary<ElementId, List<Element>>();
            try
            {
                List<RevitLinkInstance> slabLinkInstances = method.findLinkInstWithSlab(doc);
                Transform totalTransform = null;
                Transform inverseTransform = null;
                if (slabLinkInstances.Count != 0)
                {
                    foreach (RevitLinkInstance linkedInst in slabLinkInstances)
                    {
                        if (!hangerSlabDict.Keys.Contains(inst.Id))
                        {
                            totalTransform = linkedInst.GetTotalTransform();
                            inverseTransform = totalTransform.Inverse;
                            FilteredElementCollector collectorSB = method.getAllLinkedSlab(linkedInst.GetLinkDocument());
                            foreach (Element e in collectorSB)
                            {
                                Solid slabSolid = singleSolidFromElement(e);
                                if (slabSolid == null) continue;
                                //針對樓板進行變換
                                slabSolid = SolidUtils.CreateTransformed(slabSolid, totalTransform);
                                //BoundingBoxXYZ slabBox = e.get_BoundingBox(null);
                                BoundingBoxXYZ slabBox = slabSolid.GetBoundingBox();
                                XYZ solidCenter = slabSolid.ComputeCentroid();
                                Transform newTrans = Transform.Identity;
                                newTrans.Origin = solidCenter;
                                Outline outLine = new Outline(newTrans.OfPoint(slabBox.Min), newTrans.OfPoint(slabBox.Max));
                                //快慢速過濾器結合->和樓板進行碰撞
                                BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                                ElementIntersectsSolidFilter elementIntersectsSolidFilter = new ElementIntersectsSolidFilter(slabSolid);
                                ICollection<ElementId> hangerToUse = new List<ElementId>() { inst.Id };
                                List<Element> slabToUse = new FilteredElementCollector(doc, hangerToUse).WherePasses(boundingBoxIntersectsFilter)/*.WherePasses(elementIntersectsSolidFilter)*/.ToList();
                                if (slabToUse.Count() != 0 && !hangerSlabDict.Keys.Contains(inst.Id))
                                {
                                    hangerSlabDict.Add(inst.Id, slabToUse);
                                }
                                else if (slabToUse.Count() != 0 && hangerSlabDict.Keys.Contains(inst.Id))
                                {
                                    foreach (Element ee in slabToUse)
                                    {
                                        hangerSlabDict[inst.Id].Add(ee);
                                    }
                                }
                                else if (!hangerSlabDict.Keys.Contains(inst.Id))
                                {
                                    List<Element> tempList = new List<Element>();
                                    hangerSlabDict.Add(inst.Id, tempList);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("無法判斷吊架與樓板的關係!");
            }
            return hangerSlabDict;
        }
        //public FilteredElementCollector getAllLinkedSlab(Document linkedDoc)
        //{
        //    FilteredElementCollector slabCollector = new FilteredElementCollector(linkedDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
        //    return slabCollector;
        //}
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
            IList<Solid> fromElement = GetTargetSolids(inputElement);
            int solidCount = fromElement.Count;
            Solid solidResult = null;
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

        public double sortLevelbyHeight(Element element)
        {
            Level tempLevel = element as Level;
            double levelHeight = tempLevel.Elevation;
            return levelHeight;
        }
        public double getPipeDia(Element element)
        {
#if RELEASE2019
            DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
            ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
            Parameter targetPara = null;
            //Pipe
            if (element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
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
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            }
            //利用管徑(doubleType)來判斷
            double convertUnit = UnitUtils.ConvertFromInternalUnits(targetPara.AsDouble(), unitType);
            return targetPara.AsDouble();
            //return convertUnit;
        }
        public double getOutterDia(Element element)
        {
#if RELEASE2019
            DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
            ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
            Parameter targetPara = null;
            //Pipe
            if (element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            }
            //Conduit
            else if (element.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM);
            }
            //Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            }
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            }
            //利用管徑(doubleType)來判斷
            double convertUnit = UnitUtils.ConvertFromInternalUnits(targetPara.AsDouble(), unitType);
            //return convertUnit;
            return targetPara.AsDouble();
        }
        //public bool checkPara(Element elem, string paraName)
        //{
        //    bool result = false;
        //    foreach (Parameter parameter in elem.Parameters)
        //    {
        //        Parameter val = parameter;
        //        if (val.Definition.Name == paraName)
        //        {
        //            result = true;
        //        }
        //    }
        //    return result;
        //}
        public static ModelLine CreateModelLine(
/*Autodesk.Revit.ApplicationServices.Application app,*/Document doc,
  XYZ p,
  XYZ q)
        {
            if (p.DistanceTo(q) < Util.MinLineLength) return null;

            XYZ v = q - p;

            double dxy = Math.Abs(v.X) + Math.Abs(v.Y);

            XYZ w = (dxy > Util.TolPointOnPlane)
              ? XYZ.BasisZ
              : XYZ.BasisY;

            XYZ norm = v.CrossProduct(w).Normalize();

            Plane plane = Plane.CreateByNormalAndOrigin(norm, p);

            //Autodesk.Revit.Creation.Document creDoc
            //  = app.ActiveDocument.Create;

            SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
            Line tempLine = Line.CreateBound(p, q);
            ModelLine mdlLine = doc.Create.NewModelCurve(tempLine, sketchPlane) as ModelLine;
            //ModelLine mdlLine = Line.CreateBound(newTargetPt, upperPoint);


            //return creDoc.NewModelCurve(
            //  app.Create.NewLine(p, q, true),
            //  sketchPlane) as ModelLine;
            return mdlLine;
        }
    }
}
public class ElementWithTrans
{
    public Element pipe { get; set; }
    public Curve locateCrv { get; set; }
    public Transform trans { get; set; }
}
