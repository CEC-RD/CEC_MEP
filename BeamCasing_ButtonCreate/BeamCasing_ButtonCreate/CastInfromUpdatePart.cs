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
using System.Text;
using System.IO;
using System.Threading;
#endregion

namespace BeamCasing_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CastInfromUpdatePart : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public static string errorOutput = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw.Reset();//碼表歸零
            sw.Start();//碼表開始計時
            try
            {
                //步驟上進行如下，先挑出所有的RC樑，檢查RC樑中是否有ST樑，
                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                //找到所有穿樑套管元件
                List<FamilyInstance> famList = findTargetElements(doc);
                using (CastUpdateProgressUI progressView = new CastUpdateProgressUI(famList.Count))
                {
                    List<string> usefulParaName = new List<string> { "BTOP", "BCOP", "BBOP", "TTOP", "TCOP", "TBOP",
                    "【原則】上邊距", "【原則】下邊距", "【原則】尺寸檢討", "【原則】是否穿樑", "【原則】邊距檢討", "【原則】樑端檢討",
                    "干涉管數量", "系統別", "貫穿樑尺寸", "貫穿樑材料", "貫穿樑編號","開口長","不更新系統別"};
                    List<double> settingPara = new List<double>() {
                    BeamCast_Settings.Default.cD1_Ratio,
                    BeamCast_Settings.Default.cP1_Ratio,
                    BeamCast_Settings.Default.cMax1_Ratio,
                    BeamCast_Settings.Default.rD1_Ratio,
                    BeamCast_Settings.Default.rP1_Ratio,
                    BeamCast_Settings.Default.rMax1_RatioD,
                    BeamCast_Settings.Default.rMax1_RatioW,
                    BeamCast_Settings.Default.cD2_Ratio,
                    BeamCast_Settings.Default.cP2_Ratio,
                    BeamCast_Settings.Default.cMax2_Ratio,
                    BeamCast_Settings.Default.rD2_Ratio,
                    BeamCast_Settings.Default.rP2_Ratio,
                    BeamCast_Settings.Default.rMax2_RatioD,
                    BeamCast_Settings.Default.rMax2_RatioW};

                    List<double> settingPara2 = new List<double>()
                { BeamCast_Settings.Default.cP1_Min,BeamCast_Settings.Default.cMax1_Max,BeamCast_Settings.Default.rP1_Min,
                    BeamCast_Settings.Default.cP2_Min,BeamCast_Settings.Default.cMax2_Max,BeamCast_Settings.Default.rP2_Min
                };

                    //檢查設定值是否正確
                    foreach (double d in settingPara)
                    {
                        if (d == 0)
                        {
                            message = "穿樑原則的設定值有誤(非選填數值不可為0)，請重新設定後再次執行";
                            return Result.Failed;
                        }
                    }
                    for (int i = 0; i < settingPara2.Count; i++)
                    {
                        if (settingPara2[i] == null)
                        {
                            settingPara2[i] = 0;
                        }
                    }

                    //檢查參數
                    foreach (FamilyInstance famInst in famList)
                    {
                        foreach (string item in usefulParaName)
                        {
                            if (!checkPara(famInst, item))
                            {
                                MessageBox.Show($"執行失敗，請檢查{famInst.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                                return Result.Failed;
                            }
                        }
                    }

                    //更新穿樑套管參數
                    using (Transaction trans = new Transaction(doc))
                    {
                        //Dictionary<ElementId, List<Element>> castDict = getCastBeamDict(doc);
                        Dictionary<ElementId, BeamOpening> castDict_New = getCastBeamDict(doc);
                        trans.Start("更新關樑套管參數");
                        foreach (ElementId tempId in castDict_New.Keys)
                        {
                            if (castDict_New[tempId].isSC == true)
                            {
                                updateCastInst(doc.GetElement(tempId), castDict_New[tempId].scBeamsList.First());
                                updateCastMaterial(doc.GetElement(tempId), castDict_New[tempId].rcBeamsList.First());
                            }
                            else if (castDict_New[tempId].isSC != true)
                            {
                                updateCastInst(doc.GetElement(tempId), castDict_New[tempId].rcBeamsList.First());
                                updateCastMaterial(doc.GetElement(tempId), castDict_New[tempId].rcBeamsList.First());
                            }
                            //更新穿樑套管資訊&內容
                            updateCastContent(doc, doc.GetElement(tempId));
                            //如果只有一支，以該支樑為準
                            modifyCastLen(doc.GetElement(tempId), castDict_New[tempId].rcBeamsList[0]);

                            //以外參進行單位的更新進度顯示
                            if (progressView.Update()) break;
                        }

                        //檢查所有的實做套管，如果不再Dictionary中，則表示其沒有穿樑，「【原則】是否穿樑」應該設為不符合
                        List<FamilyInstance> famListUpdate = findTargetElements(doc);
                        foreach (FamilyInstance inst in famListUpdate)
                        {
                            //if (!castDict.Keys.Contains(inst.Id))
                            if (!castDict_New.Keys.Contains(inst.Id))
                            {
                                inst.LookupParameter("【原則】是否穿樑").Set("不符合");
                                inst.LookupParameter("貫穿樑尺寸").Set("無尺寸");
                                inst.LookupParameter("貫穿樑編號").Set("無編號");
                                inst.LookupParameter("干涉管數量").Set(0);
                                inst.LookupParameter("系統別").Set("SP");
                            }
                        }
                        trans.Commit();
                    }
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            sw.Stop();//碼錶停止
            double sec = Math.Round(sw.Elapsed.TotalMilliseconds / 1000, 2);
            string output = $"穿樑套管資訊更新完成，共花費 {sec} 秒\n";
            MessageBox.Show(output + errorOutput);
            errorOutput = "";
            return Result.Succeeded;
        }
        public double calculateSolidVol(Element inst, Element linkedBeam, Transform toTrans)
        {
            //計算套管與牆之間的交集量體，並回傳其體積大小
            double vol = 0.0;
            Solid instSolid = singleSolidFromElement(inst);
            Solid linkedWallSolid = singleSolidFromElement(linkedBeam);
            linkedWallSolid = SolidUtils.CreateTransformed(linkedWallSolid, toTrans);
            Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(instSolid, linkedWallSolid, BooleanOperationsType.Intersect);
            if (Math.Abs(interSolid.Volume) > 0.000001)
            {
                vol = interSolid.Volume;
            }
            return vol;
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
        public List<RevitLinkInstance> getLinkedInstances(Document doc, string materialName)
        {
            List<RevitLinkInstance> targetLinkedInstances = new List<RevitLinkInstance>();
            ElementCategoryFilter linkedFileFilter = new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks);
            FilteredElementCollector linkedFileCollector = new FilteredElementCollector(doc).WherePasses(linkedFileFilter).WhereElementIsNotElementType();
            try
            {
                if (linkedFileCollector.Count() > 0)
                {
                    foreach (RevitLinkInstance linkedInst in linkedFileCollector)
                    {
                        Document linkDoc = linkedInst.GetLinkDocument();
                        bool isLoaded = RevitLinkType.IsLoaded(doc, linkedInst.GetTypeId());
                        //if (linkDoc == null /*|| !linkedInst.IsValidObject*/ || !isLoaded) continue;
                        if (linkDoc != null && isLoaded)
                        {
                            FilteredElementCollector linkedBeams = new FilteredElementCollector(linkDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                            if (linkedBeams.Count() == 0) continue;
                            foreach (Element e in linkedBeams)
                            {
                                FamilyInstance beamInstance = e as FamilyInstance;
                                if (beamInstance.StructuralMaterialType.ToString() == materialName && !targetLinkedInstances.Contains(linkedInst))
                                {
                                    //RCfile = linkDoc;
                                    targetLinkedInstances.Add(linkedInst);
                                }
                            }
                        }
                    }
                }
                else if (linkedFileCollector.Count() == 0)
                {
                    MessageBox.Show("模型中沒有實做的外參檔案");
                }
            }
            catch
            {
                MessageBox.Show("請檢查外參連結是否為載入或有問題!");
            }
            return targetLinkedInstances;
        }
        public RevitLinkInstance getTargetLinkedInstance(Document doc, string linkTilte)
        {
            RevitLinkInstance targetLinkInstance = null;
            try
            {
                ElementCategoryFilter linkedFileFilter = new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks);
                FilteredElementCollector linkedFileCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType();
                if (linkedFileCollector.Count() > 0)
                {
                    foreach (RevitLinkInstance linkedInst in linkedFileCollector)
                    {
                        //if (linkedInst.GetLinkDocument().Title == linkTilte)
                        if (linkedInst.Name.Contains(linkTilte))
                        {
                            targetLinkInstance = linkedInst;
                            break;
                        }
                    }
                }
                else if (targetLinkInstance == null)
                {
                    MessageBox.Show("未找到對應的實做Revit外參檔!!");
                }
            }
            catch
            {
                MessageBox.Show("找尋連結實體發生問題!");
            }
            return targetLinkInstance;
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
        public FilteredElementCollector getAllLinkedBeam(Document linkedDoc)
        {
            FilteredElementCollector beamCollector = new FilteredElementCollector(linkedDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
            beamCollector.WhereElementIsNotElementType();
            return beamCollector;
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
        public List<FamilyInstance> findTargetElements(Document doc)
        {
            //RC套管跟SC開口的內部名稱是不同的
            string internalNameST = "CEC-穿樑開口";
            string internalNameRC = "CEC-穿樑套管";
            List<FamilyInstance> castInstances = new List<FamilyInstance>();
            try
            {
                FilteredElementCollector coll = new FilteredElementCollector(doc, doc.ActiveView.Id);
                ElementCategoryFilter castCate_Filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementClassFilter castInst_Filter = new ElementClassFilter(typeof(Instance));
                LogicalAndFilter andFilter = new LogicalAndFilter(castCate_Filter, castInst_Filter);
                coll.WherePasses(andFilter).WhereElementIsNotElementType().ToElements(); //找出模型中實做的穿樑套管元件
                if (coll != null)
                {
                    foreach (FamilyInstance e in coll)
                    {
                        Parameter p = e.Symbol.LookupParameter("API識別名稱");
                        if (p != null && p.AsString().Contains(internalNameST))
                        {
                            castInstances.Add(e);
                        }
                        else if (p != null && p.AsString().Contains(internalNameRC))
                        {
                            castInstances.Add(e);
                        }
                    }
                }
                else if (castInstances.Count() == 0)
                {
                    {
                        MessageBox.Show("尚未匯入套管元件，或模型中沒有實做的套管元件");
                    }
                }
            }
            catch
            {
                //MessageBox.Show("蒐集套管發生問題!");
            }
            return castInstances;
        }
        public List<FamilyInstance> findLinkTargetElements(Document doc)
        {
            List<RevitLinkInstance> linkInstanceList = getMEPLinkInstance(doc);
            List<FamilyInstance> targetList = new List<FamilyInstance>();
            foreach (RevitLinkInstance revitLink in linkInstanceList)
            {
                Document linkDoc = revitLink.GetLinkDocument();
                Transform trans = revitLink.GetTotalTransform();
                FilteredElementCollector elementCollector = new FilteredElementCollector(linkDoc);
                List<FamilyInstance> tempList = findTargetElements(linkDoc);
                if (tempList.Count() > 0)
                {
                    foreach (FamilyInstance inst in tempList)
                    {
                        targetList.Add(inst);
                    }
                }
            }
            return targetList;
        }
        //可以判斷樑中樑的程式-->用Dictionary裝穿樑套管以及與之對應的樑
        //因為樑與套管正常來說應該是一對一的關係，強制取得它們的關係(用套管ID反查干涉的樑)
        //public Dictionary<ElementId, List<Element>> getCastBeamDict(Document doc)
        public Dictionary<ElementId, BeamOpening> getCastBeamDict(Document doc)
        {
            Dictionary<ElementId, BeamOpening> castBeamDict_New = new Dictionary<ElementId, BeamOpening>();
            FamilyInstance instance = null;
            try
            {
                //要確定鋼構模型和RC結構模型是否有確實拆開
                List<FamilyInstance> familyInstances = findTargetElements(doc);
                List<RevitLinkInstance> SCLinkedInstance = getLinkedInstances(doc, "Steel");
                List<RevitLinkInstance> RCLinkedInstance = getLinkedInstances(doc, "Concrete");
                Transform totalTransform = null;
                Transform inverseTransform = null;

                if (RCLinkedInstance.Count != 0 || SCLinkedInstance.Count != 0)
                {
                    foreach (FamilyInstance inst in familyInstances)
                    {
                        bool toAdd = false;
                        BeamOpening newCast = new BeamOpening();
                        //將RC和ST的檔案分別與inst去做碰撞，取得有用效的樑，再用dictionary的key值判斷套管是否已經存在字典之中，有才進行執行
                        foreach (RevitLinkInstance SClinkedInst in SCLinkedInstance)
                        {
                            totalTransform = SClinkedInst.GetTotalTransform();
                            inverseTransform = totalTransform.Inverse;
                            FilteredElementCollector collectorSC = getAllLinkedBeam(SClinkedInst.GetLinkDocument());
                            Solid castSolid = singleSolidFromElement(inst);
                            if (castSolid == null) continue;

                            //座標變換，因為抓法的關係，要轉換成外參檔「原本」的座標
                            castSolid = SolidUtils.CreateTransformed(castSolid, inverseTransform);
                            BoundingBoxXYZ solidBounding = castSolid.GetBoundingBox();
                            XYZ solidCenter = castSolid.ComputeCentroid();
                            Transform newTrans = Transform.Identity;
                            newTrans.Origin = solidCenter;
                            Outline outLine = new Outline(newTrans.OfPoint(solidBounding.Min), newTrans.OfPoint(solidBounding.Max));

                            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                            ElementIntersectsSolidFilter elementIntersectsSolidFilter = new ElementIntersectsSolidFilter(castSolid);
                            //if (collectorSC.Count() == 0) continue;
                            collectorSC.WherePasses(boundingBoxIntersectsFilter).WherePasses(elementIntersectsSolidFilter);
                            List<Element> tempList = collectorSC.ToList();
                            if (tempList.Count > 0) toAdd = true;
                            if (tempList.Count > 0)
                            {
                                newCast.scLinkInstance = SClinkedInst;
                                newCast.isSC = true;
                                newCast.scBeamsList = tempList;
                            }
                        }
                        //和SC沒撞出東西，再和RC撞，如果被撞到的套管ID已在字典Key裡，則略過(後來不能略過，因為需要RC外框來更新套管長度)
                        foreach (RevitLinkInstance RClinkedInst in RCLinkedInstance)
                        {
                            totalTransform = RClinkedInst.GetTotalTransform();
                            inverseTransform = totalTransform.Inverse;
                            FilteredElementCollector collectorRC = getAllLinkedBeam(RClinkedInst.GetLinkDocument());
                            Solid castSolid = singleSolidFromElement(inst);
                            if (castSolid == null) continue;

                            //座標變換，因為抓法的關係，要轉換成外參檔「原本」的座標
                            castSolid = SolidUtils.CreateTransformed(castSolid, inverseTransform);
                            BoundingBoxXYZ solidBounding = castSolid.GetBoundingBox();
                            XYZ solidCenter = castSolid.ComputeCentroid();
                            Transform newTrans = Transform.Identity;
                            newTrans.Origin = solidCenter;
                            Outline outLine = new Outline(newTrans.OfPoint(solidBounding.Min), newTrans.OfPoint(solidBounding.Max));

                            #region 尚未轉換之前的抓法，如果座標有變換可能會有問題
                            //BoundingBoxXYZ castBounding = inst.get_BoundingBox(null);
                            //Transform t = castBounding.Transform;
                            //Outline outLine = new Outline(t.OfPoint(castBounding.Min), t.OfPoint(castBounding.Max));
                            ////Outline outLine = new Outline(castBounding.Min, castBounding.Max);
                            #endregion
                            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                            ElementIntersectsSolidFilter elementIntersectsSolidFilter = new ElementIntersectsSolidFilter(castSolid);
                            collectorRC.WherePasses(boundingBoxIntersectsFilter).WherePasses(elementIntersectsSolidFilter);
                            //if (collectorRC.Count() == 0) continue;
                            //MessageBox.Show($"{inst.Id}的套管沒有和任何樑干涉");
                            ////List<Element> tempList = collectorSC.ToList();
                            List<Element> tempList = collectorRC.OrderByDescending(x => calculateSolidVol(doc.GetElement(inst.Id), x, totalTransform)).ToList();
                            //MessageBox.Show("YA");
                            //inst.LookupParameter("備註").Set(collectorRC.Count().ToString());
                            //List<Element> targetList = new List<Element>() { tempList.First() };
                            if (tempList.Count > 0) toAdd = true;
                            if (tempList.Count > 0)
                            {
                                if (newCast.isSC == true) newCast.isSRC = true;
                                newCast.rcLinkInstance = RClinkedInst;
                                newCast.rcBeamsList = tempList;
                            }
                        }
                        if (toAdd)
                        {
                            castBeamDict_New.Add(inst.Id, newCast);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show($"無法判斷套管{instance.Id}與樑的關係");
                //MessageBox.Show("無法判斷套管與樑的關係!");
            }
            return castBeamDict_New;

        }
        //BoundingBox外擴
        public BoundingBoxXYZ ExtandBounding(BoundingBoxXYZ bounding, double extensionValue)
        {
            extensionValue = UnitUtils.ConvertToInternalUnits(extensionValue, unitType);
            bounding.Max = new XYZ(bounding.Max.X + extensionValue, bounding.Max.Y + extensionValue, bounding.Max.Z + extensionValue);
            bounding.Min = new XYZ(bounding.Min.X - extensionValue, bounding.Min.Y - extensionValue, bounding.Min.Z - extensionValue);
            return bounding;
        }
        public bool checkGrider(Element elem)
        {
            bool result = false;
            try
            {
                Document doc = elem.Document;
                FilteredElementCollector tempCollector = new FilteredElementCollector(doc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralColumns);
                BoundingBoxXYZ checkBounding = elem.get_BoundingBox(null);
                Autodesk.Revit.DB.Transform t1 = checkBounding.Transform;
                Outline outline1 = new Outline(t1.OfPoint(checkBounding.Min), t1.OfPoint(checkBounding.Max));
                BoundingBoxIntersectsFilter boundingBoxIntersectsFilter1 = new BoundingBoxIntersectsFilter(outline1, 0.1);
                //ElementIntersectsSolidFilter elementIntersectsSolidFilter1 = new ElementIntersectsSolidFilter(RCSolid);
                tempCollector.WherePasses(boundingBoxIntersectsFilter1);
                if (tempCollector.Count() > 0)
                {
                    result = true;
                }
                else if (tempCollector.Count() == 0)
                {
                    result = false;
                }
            }
            catch
            {
                MessageBox.Show("判斷大小樑失敗!");
            }
            return result;
        }
        public double getCastWidth(Element elem)
        {
            double targetWidth = 0.0;
            Parameter widthPara = null;
            FamilyInstance inst = elem as FamilyInstance;
            string inernalName = inst.Symbol.LookupParameter("API識別名稱").AsString();
            if (inernalName == null)
            {
                MessageBox.Show("請檢查目標元件的API識別名稱是否遭到修改");
            }
            else if (inernalName.Contains("圓"))
            {
                widthPara = inst.Symbol.LookupParameter("管外直徑");
                if (widthPara == null) MessageBox.Show($"請檢查{inst.Symbol.FamilyName}中是否缺少「管外直徑」參數欄位");
                targetWidth = widthPara.AsDouble();
            }
            else if (inernalName.Contains("方"))
            {
                widthPara = inst.LookupParameter("開口寬");
                if (widthPara == null) MessageBox.Show($"請檢查{inst.Symbol.FamilyName}中是否缺少「開口寬」參數欄位");
                targetWidth = widthPara.AsDouble();
            }
            return targetWidth;
        }
        public double getCastHeight(Element elem)
        {
            double targetWidth = 0.0;
            Parameter widthPara = null;
            FamilyInstance inst = elem as FamilyInstance;
            string inernalName = inst.Symbol.LookupParameter("API識別名稱").AsString();
            if (inernalName == null)
            {
                MessageBox.Show("請檢查目標元件的API識別名稱是否遭到修改");
            }
            else if (inernalName.Contains("圓"))
            {
                widthPara = inst.Symbol.LookupParameter("管外直徑");
                if (widthPara == null) MessageBox.Show($"請檢查{inst.Symbol.FamilyName}中是否缺少「管外直徑」參數欄位");
                targetWidth = widthPara.AsDouble();
            }
            else if (inernalName.Contains("方"))
            {
                widthPara = inst.LookupParameter("開口高");
                if (widthPara == null) MessageBox.Show($"請檢查{inst.Symbol.FamilyName}中是否缺少「H」參數欄位");
                targetWidth = widthPara.AsDouble();
            }
            return targetWidth;
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
        public Element modifyCastLen(Element elem, Element linkedBeam)
        {
            //這個功能應該還要視樑為RC還是SRC去決定寬度值
            //先利用linkedBeam反找revitLinkedInstance
            FamilyInstance updateCast = null;
            try
            {
                Document document = elem.Document;
                RevitLinkInstance targetLink = getTargetLinkedInstance(document, linkedBeam.Document.Title);
                Transform linkedInstTrans = targetLink.GetTotalTransform();

                //計算偏移值&設定長度
                LocationCurve beamLocate = linkedBeam.Location as LocationCurve;
                Curve beamCurve = beamLocate.Curve;
                XYZ startPoint = beamCurve.GetEndPoint(0);
                XYZ endPoint = beamCurve.GetEndPoint(1);
                startPoint = TransformPoint(startPoint, linkedInstTrans);
                endPoint = TransformPoint(endPoint, linkedInstTrans);
                endPoint = new XYZ(endPoint.X, endPoint.Y, startPoint.Z);
                Line tempCrv = Line.CreateBound(startPoint, endPoint);

                LocationPoint castLocate = elem.Location as LocationPoint;
                XYZ castPt = castLocate.Point;
                XYZ tempPt = new XYZ(castPt.X, castPt.Y, startPoint.Z);
                IntersectionResult intersectResult = tempCrv.Project(castPt);
                XYZ targetPoint = intersectResult.XYZPoint;
                targetPoint = new XYZ(targetPoint.X, targetPoint.Y, castPt.Z);
                XYZ positionChange = targetPoint - castPt;
                double castLength = getBeamWidth(linkedBeam) + 2 / 30.48;
                //castLength = UnitUtils.ConvertToInternalUnits(castLength, unitType);

                FamilyInstance inst = elem as FamilyInstance;
                Parameter instLenPara = inst.LookupParameter("開口長");
                double beamWidth = getBeamWidth(linkedBeam);
                //先調整套管位置
                if (elem.Pinned == true)
                {
                    errorOutput += $"注意，ID為 {elem.Id} 的套管被固定無法移動，請檢查是否為有意不想移動之元素!\n";
                }
                else if (!castPt.IsAlmostEqualTo(targetPoint) && elem.Pinned == false)
                {
                    ElementTransformUtils.MoveElement(document, inst.Id, positionChange);
                    //再調整套管長度
                    if (instLenPara.AsDouble() <= beamWidth)
                    {
                        instLenPara.Set(castLength);
                    }
                }
                updateCast = inst;
            }
            catch
            {
                errorOutput += $"更新套管長度失敗，ID為 {elem.Id} 的套管無法更新長度!\n";
                //MessageBox.Show($"更新套管長度失敗，ID為{elem.Id}的套管無法更新長度!");
            }
            return updateCast;
        }
        public Element updateCastInst(Element elem, Element linkedBeam)
        {
            FamilyInstance updateCast = null;
            try
            {
                FamilyInstance inst = elem as FamilyInstance;
                Document document = elem.Document;
                RevitLinkInstance targetLink = getTargetLinkedInstance(document, linkedBeam.Document.Title);
                //MessageBox.Show(targetLink.Name);
                Transform linkedInstTrans = targetLink.GetTotalTransform();
                Solid beamSolid = singleSolidFromElement(linkedBeam);
                beamSolid = SolidUtils.CreateTransformed(beamSolid, linkedInstTrans);

                SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();
                double beamHeight = 0.0;
                if (null != beamSolid)
                {
                    LocationPoint instLocate = inst.Location as LocationPoint;
                    double inst_CenterZ = (inst.get_BoundingBox(null).Max.Z + inst.get_BoundingBox(null).Min.Z) / 2;
                    XYZ instPt = instLocate.Point;
                    double normal_BeamHeight = UnitUtils.ConvertToInternalUnits(1500, unitType);
                    XYZ inst_Up = new XYZ(instPt.X, instPt.Y, instPt.Z + normal_BeamHeight);
                    XYZ inst_Dn = new XYZ(instPt.X, instPt.Y, instPt.Z - normal_BeamHeight);
                    Curve instVerticalCrv = Autodesk.Revit.DB.Line.CreateBound(inst_Dn, inst_Up);
                    //這邊用solid是因為怕有斜樑需要開口的問題，但斜樑的結構應力應該已經蠻集中的，不可以再開口
                    SolidCurveIntersection intersection = beamSolid.IntersectWithCurve(instVerticalCrv, options);
                    int intersectCount = intersection.SegmentCount;
                    //針對有切割到的實體去做計算六個參數
                    if (intersectCount > 0)
                    {
                        string instInternalName = inst.Symbol.LookupParameter("API識別名稱").AsString();
                        //針對有交集的實體去做計算
                        inst.LookupParameter("【原則】是否穿樑").Set("OK");
                        //計算TOP、BOP等六個參數
                        LocationPoint cast_Locate = inst.Location as LocationPoint;
                        XYZ LocationPt = cast_Locate.Point;
                        XYZ cast_Max = inst.get_BoundingBox(null).Max;
                        XYZ cast_Min = inst.get_BoundingBox(null).Min;
                        Curve castIntersect_Crv = intersection.GetCurveSegment(0);
                        XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                        XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);
                        double castCenter_Z = (cast_Max.Z + cast_Min.Z) / 2;
                        double TTOP_update = intersect_UP.Z - cast_Max.Z;
                        double BTOP_update = cast_Max.Z - intersect_DN.Z;
                        double TCOP_update = intersect_UP.Z - castCenter_Z;
                        double BCOP_update = castCenter_Z - intersect_DN.Z;
                        double TBOP_update = intersect_UP.Z - cast_Min.Z;
                        double BBOP_update = cast_Min.Z - intersect_DN.Z;
                        double TTOP_orgin = inst.LookupParameter("TTOP").AsDouble();
                        double BBOP_orgin = inst.LookupParameter("BBOP").AsDouble();
                        beamHeight = intersect_UP.Z - intersect_DN.Z;
                        double test = beamHeight * 30.48;
                        double castHeight = cast_Max.Z - cast_Min.Z;
                        double TTOP_Check = Math.Round(UnitUtils.ConvertFromInternalUnits(TTOP_update, unitType), 1);
                        double TTOP_orginCheck = Math.Round(UnitUtils.ConvertFromInternalUnits(TTOP_orgin, unitType), 1);
                        double BBOP_Check = Math.Round(UnitUtils.ConvertFromInternalUnits(BBOP_update, unitType), 1);
                        double BBOP_orginCheck = Math.Round(UnitUtils.ConvertFromInternalUnits(BBOP_orgin, unitType), 1);

                        if (TTOP_Check != TTOP_orginCheck || BBOP_Check != BBOP_orginCheck)
                        {
                            inst.LookupParameter("TTOP").Set(TTOP_update);
                            inst.LookupParameter("BTOP").Set(BTOP_update);
                            inst.LookupParameter("TCOP").Set(TCOP_update);
                            inst.LookupParameter("BCOP").Set(BCOP_update);
                            inst.LookupParameter("TBOP").Set(TBOP_update);
                            inst.LookupParameter("BBOP").Set(BBOP_update);
                        }
                        //寫入樑編號與樑尺寸，先檢查樑是否有缺參數
                        string beamSIze = linkedBeam.LookupParameter("類型").AsValueString();
                        Parameter beamNum = linkedBeam.LookupParameter("編號");
                        Parameter instBeamNum = inst.LookupParameter("貫穿樑編號");
                        Parameter instBeamSize = inst.LookupParameter("貫穿樑尺寸");

                        if (beamNum != null)
                        {
                            instBeamNum.Set(beamNum.AsString());
                        }
                        else
                        {
                            instBeamNum.Set("無編號");
                        }
                        if (beamSIze != null)
                        {
                            instBeamSize.Set(beamSIze);
                        }
                        else
                        {
                            instBeamSize.Set("無尺寸");
                        }

                        //設定檢核結果，在此之前感覺需要做一個找到元件寬度的程式
                        double protectDistCheck = 0.0; //確認套管是否與保護層過近
                        double sizeMaxCheck = 0.0; //確認套管是否過大
                        double sizeMaxCheckW = 0.0;
                        double sizeMaxCheckD = 0.0;
                        double endDistCheck = 0.0; //確認套管是否與樑末端過近
                        bool isGrider = checkGrider(linkedBeam); //確認是否為大小樑，以此作為參數更新的依據

                        //依照是否為大小樑，更新參數依據
                        double C_distRatio = 0.0, C_protectRatio = 0.0, C_protectMin = 0.0, C_sizeRatio = 0.0, C_sizeMax = 0.0; ;
                        double R_distRatio = 0.0, R_protectRatio = 0.0, R_protectMin = 0.0, R_sizeRatioD = 0.0, R_sizeRatioW = 0.0;
                        if (isGrider)
                        {
                            C_distRatio = BeamCast_Settings.Default.cD1_Ratio;
                            C_protectRatio = BeamCast_Settings.Default.cP1_Ratio;
                            C_protectMin = BeamCast_Settings.Default.cP1_Min;
                            C_sizeRatio = BeamCast_Settings.Default.cMax1_Ratio;
                            C_sizeMax = BeamCast_Settings.Default.cMax1_Max;
                            R_distRatio = BeamCast_Settings.Default.rD1_Ratio;
                            R_protectRatio = BeamCast_Settings.Default.rP1_Ratio;
                            R_protectMin = BeamCast_Settings.Default.rP1_Min;
                            R_sizeRatioD = BeamCast_Settings.Default.rMax1_RatioD;
                            R_sizeRatioW = BeamCast_Settings.Default.rMax1_RatioW;
                        }
                        else if (!isGrider)
                        {
                            C_distRatio = BeamCast_Settings.Default.cD2_Ratio;
                            C_protectRatio = BeamCast_Settings.Default.cP2_Ratio;
                            C_protectMin = BeamCast_Settings.Default.cP2_Min;
                            C_sizeRatio = BeamCast_Settings.Default.cMax2_Ratio;
                            C_sizeMax = BeamCast_Settings.Default.cMax2_Max;
                            R_distRatio = BeamCast_Settings.Default.rD2_Ratio;
                            R_protectRatio = BeamCast_Settings.Default.rP2_Ratio;
                            R_protectMin = BeamCast_Settings.Default.rP2_Min;
                            R_sizeRatioD = BeamCast_Settings.Default.rMax2_RatioD;
                            R_sizeRatioW = BeamCast_Settings.Default.rMax2_RatioW;
                        }
                        List<double> parameter_Checklist = new List<double> { C_distRatio, C_protectRatio, C_sizeRatio, R_distRatio, R_protectRatio, R_sizeRatioD, R_sizeRatioW };
                        List<double> parameter_Checklist2 = new List<double> { C_protectMin, C_sizeMax, R_protectMin };

                        bool isCircleCast = instInternalName.Contains("圓");
                        //前面已經檢查過是否為大小樑，在此檢查是方孔還圓孔，比例係數不同
                        //如果套管為圓形
                        if (isCircleCast)
                        {
                            //上下邊距警告值
                            protectDistCheck = C_protectRatio * beamHeight;
                            double tempProtectValue = UnitUtils.ConvertToInternalUnits(C_protectMin, unitType);
                            if (tempProtectValue > protectDistCheck) protectDistCheck = tempProtectValue;

                            //最大尺寸警告值
                            sizeMaxCheck = C_sizeRatio * beamHeight;
                            double tempSizeValue = UnitUtils.ConvertToInternalUnits(C_sizeMax, unitType);
                            if (tempSizeValue < sizeMaxCheck && tempSizeValue != 0) sizeMaxCheck = tempSizeValue;

                            //樑兩端警告值
                            endDistCheck = C_distRatio * beamHeight;
                        }
                        //如果套管為方形
                        else if (!isCircleCast)
                        {
                            //上下邊距警告值
                            protectDistCheck = R_protectRatio * beamHeight;
                            double tempProtectValue = UnitUtils.ConvertToInternalUnits(R_protectMin, unitType);
                            if (tempProtectValue < protectDistCheck) protectDistCheck = tempProtectValue;

                            //最大尺寸警告值
                            sizeMaxCheckW = R_sizeRatioW * beamHeight;
                            sizeMaxCheckD = R_sizeRatioD * beamHeight;

                            //樑兩端警告值
                            endDistCheck = R_distRatio * beamHeight;
                        }
                        //檢查是否穿樑
                        List<double> updateParas = new List<double> { TTOP_update, BTOP_update, TCOP_update, BCOP_update, TBOP_update, BBOP_update };
                        foreach (double d in updateParas)
                        {
                            if (d < 0)
                            {
                                inst.LookupParameter("【原則】是否穿樑").Set("不符合");
                                instBeamNum.Set("無編號");
                                instBeamSize.Set("無尺寸");
                            }
                        }
                        //檢查是否過大
                        Parameter sizeCheckPara = inst.LookupParameter("【原則】尺寸檢討");
                        if (isCircleCast)
                        {

                            double castSize = getCastWidth(inst);
                            if (castSize > sizeMaxCheck) sizeCheckPara.Set("不符合");
                            else sizeCheckPara.Set("OK");
                        }
                        else if (!isCircleCast)
                        {
                            double castSizeW = getCastWidth(inst);
                            double castSizeD = getCastHeight(inst);
                            if (castSizeW > sizeMaxCheckW || castSizeD > sizeMaxCheckD) sizeCheckPara.Set("不符合");
                            else sizeCheckPara.Set("OK");
                        }

                        //檢查上下部包護層
                        Parameter protectionCheckPara_UP = inst.LookupParameter("【原則】上邊距");
                        Parameter protectionCheckPara_DN = inst.LookupParameter("【原則】下邊距");
                        if (TTOP_update < protectDistCheck)
                        {
                            protectionCheckPara_UP.Set("不符合");
                        }
                        else
                        {
                            protectionCheckPara_UP.Set("OK");
                        }
                        if (BBOP_update < protectDistCheck)
                        {
                            protectionCheckPara_DN.Set("不符合");
                        }
                        else
                        {
                            protectionCheckPara_DN.Set("OK");
                        }

                        //檢查套管是否離樑的兩端過近
                        Parameter endCheckPara = inst.LookupParameter("【原則】樑端檢討");
                        LocationCurve tempLocateCrv = linkedBeam.Location as LocationCurve;
                        Curve targetCrv = tempLocateCrv.Curve;
                        XYZ tempStart = targetCrv.GetEndPoint(0);
                        XYZ tempEnd = targetCrv.GetEndPoint(1);
                        XYZ startPt = new XYZ(tempStart.X, tempStart.Y, instPt.Z);
                        XYZ endPt = new XYZ(tempEnd.X, tempEnd.Y, instPt.Z);
                        List<XYZ> points = new List<XYZ>() { startPt, endPt };
                        List<double> distLIst = new List<double>();
                        foreach (XYZ pt in points)
                        {
                            double distToBeamEnd = instPt.DistanceTo(pt);
                            distLIst.Add(distToBeamEnd);
                        }
                        if (distLIst.Min() - getCastWidth(elem) / 2 < endDistCheck)
                        {
                            endCheckPara.Set("不符合");
                        }
                        else if (distLIst.Min() - getCastHeight(elem) / 2 > endDistCheck)
                        {
                            endCheckPara.Set("OK");
                        }
                    }
                    else if (intersectCount == 0)
                    {
                        inst.LookupParameter("【原則】是否穿樑").Set("不符合");
                        inst.LookupParameter("貫穿樑編號").Set("無編號");
                        inst.LookupParameter("貫穿樑尺寸").Set("無尺寸");
                    }

                    //與其他穿樑套管之間的距離檢討
                    List<FamilyInstance> tempList = findTargetElements(elem.Document);
                    //針對外參的套管也需要抓取
                    List<FamilyInstance> linkFamList = findLinkTargetElements(elem.Document);
                    List<FamilyInstance> targetList = new List<FamilyInstance>();

                    //利用UI中的CheckBox判斷是否需要計算外參，在各子系統中需要，CSD中不用
                    if (BeamCast_Settings.Default.checkLink == true)
                    {
                        targetList = tempList.Union(linkFamList).ToList();
                    }
                    else if (BeamCast_Settings.Default.checkLink == false)
                    {
                        targetList = tempList;
                    }
                    List<double> distList = new List<double>();
                    double baseWidth = getCastWidth(elem);
                    foreach (FamilyInstance e in targetList)
                    {
                        string elemLevel = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString();
                        string eLevel = e.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString();
                        //自己跟自己不用算
                        if (e.Id == elem.Id)
                        {
                            continue;
                        }
                        //同一層的才要進行距離檢討與計算
                        //else if (e.LevelId == elem.LevelId)
                        else if (elemLevel == eLevel)
                        {
                            double targetWidth = getCastWidth(e);
                            double distCheck = baseWidth + targetWidth;
                            //不論是圓孔、方孔或兩者間測距，都要和樑深(h)比較，取小者
                            if (distCheck > beamHeight)
                            {
                                distCheck = beamHeight;
                            }
                            LocationPoint baseLocate = elem.Location as LocationPoint;
                            XYZ basePt = baseLocate.Point;
                            LocationPoint targetLocate = e.Location as LocationPoint;
                            XYZ targetPt = targetLocate.Point;
                            XYZ adjustPt = new XYZ(targetPt.X, targetPt.Y, basePt.Z);
                            double dist = basePt.DistanceTo(adjustPt);
                            if (dist / 1.5 < distCheck)
                            {
                                distList.Add(dist);
                            }
                        }
                    }
                    if (distList.Count > 0)
                    {
                        inst.LookupParameter("【原則】邊距檢討").Set("不符合");
                    }
                    else
                    {
                        inst.LookupParameter("【原則】邊距檢討").Set("OK");
                    }
                    updateCast = inst;
                }
                else if (null == beamSolid)
                {
                    MessageBox.Show($"來自{linkedBeam.Document.Title}，編號{linkedBeam.Id}的樑，無法創造一個完整的實體，因此無法更新該樑內的套管");
                }
            }
            catch
            {
                errorOutput += $"更新套管資訊失敗，ID為 {elem.Id} 的套管無法更新資訊!\n";
                //MessageBox.Show($"穿樑套管資訊更新失敗，ID為{elem.Id}的套管無法更新!");
            }
            return updateCast;
        }
        public Element updateCastContent(Document doc, Element elem)
        {
            FamilyInstance updateCast = null;
            try
            {
                List<string> systemName = new List<string>() { "E", "T", "W", "P", "F", "A", "G" };
                FamilyInstance inst = elem as FamilyInstance;
                int notUpdate = inst.LookupParameter("不更新系統別").AsInteger();
                if (notUpdate == 1) return inst;
                BuiltInCategory[] builts =
                {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_DuctCurves
            };
                List<ElementFilter> filters = new List<ElementFilter>();
                foreach (BuiltInCategory built in builts)
                {
                    ElementCategoryFilter filter = new ElementCategoryFilter(built);
                    filters.Add(filter);
                }
                LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
                FilteredElementCollector pipeCollector = new FilteredElementCollector(doc).WherePasses(categoryFilter);
                BoundingBoxXYZ castBounding = inst.get_BoundingBox(null);
                Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);
                Solid castSolid = singleSolidFromElement(inst);
                ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                pipeCollector.WherePasses(boxIntersectsFilter).WherePasses(solidFilter);
                Parameter systemType = inst.LookupParameter("系統別");

                List<RevitLinkInstance> otherMEP = getMEPLinkInstance(doc);
                List<Element> pipeCollector_final = new List<Element>();
                //將本機端蒐集到的管放進list
                foreach (Element e in pipeCollector)
                {
                    pipeCollector_final.Add(e);
                }
                //針對每一個實做的外參，蒐集管段後加入list
                foreach (RevitLinkInstance linkInst in otherMEP)
                {
                    //座標變換，因為抓法的關係，要轉換成外參檔「原本」的座標
                    Transform trans = linkInst.GetTotalTransform();
                    Transform inverseTransform = trans.Inverse;
                    Document linkdoc = linkInst.GetLinkDocument();
                    if (linkdoc != null)
                    {
                        Solid CSDsolid = SolidUtils.CreateTransformed(castSolid, inverseTransform);
                        ElementIntersectsSolidFilter solidFilter_CSD = new ElementIntersectsSolidFilter(CSDsolid);
                        BoundingBoxXYZ solidBounding = CSDsolid.GetBoundingBox();
                        XYZ solidCenter = CSDsolid.ComputeCentroid();
                        Transform newTrans = Transform.Identity;
                        newTrans.Origin = solidCenter;
                        Outline outLine = new Outline(newTrans.OfPoint(solidBounding.Min), newTrans.OfPoint(solidBounding.Max));
                        BoundingBoxIntersectsFilter boxIntersectsFilter_CSD = new BoundingBoxIntersectsFilter(outLine);
                        FilteredElementCollector CSDcollector = new FilteredElementCollector(linkdoc).WherePasses(categoryFilter).WherePasses(boxIntersectsFilter_CSD).WherePasses(solidFilter_CSD);

                        if (CSDcollector.Count() > 0)
                        {
                            foreach (Element e in CSDcollector)
                            {
                                pipeCollector_final.Add(e);
                            }
                        }
                    }
                }
                if (pipeCollector_final.Count() == 0)
                {
                    systemType.Set("SP");
                }
                inst.LookupParameter("干涉管數量").Set(pipeCollector_final.Count());
                //針對蒐集到的管去做系統別的更新，因為電管沒有系統類型，要和管分開處理
                if (pipeCollector_final.Count() == 1)
                {
                    if (pipeCollector_final.First().Category.Name == "電管")
                    {
                        systemType.Set("E");
                    }
                    else
                    {
                        Element targetPipe = pipeCollector_final.First();
                        string pipeSystem = null;
                        if (targetPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) != null)
                        {
                            pipeSystem = targetPipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                        }
                        else if (targetPipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) != null)
                        {
                            pipeSystem = targetPipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                        }
                        string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                        if (systemName.Contains(shortSystemName))
                        {
                            systemType.Set(shortSystemName);
                        }
                        else if (pipeSystem.Contains("空調"))
                        {
                            systemType.Set("A");
                        }
                        else
                        {
                            systemType.Set("SP");
                        }
                    }
                }
                //如果有共管的狀況
                else if (pipeCollector_final.Count() >= 2)
                {
                    List<int> isPipe = new List<int>();
                    List<string> shortNameList = new List<string>();
                    foreach (Element pipe in pipeCollector_final)
                    {
                        if (pipe.Category.Name == "電管")
                        {
                            isPipe.Add(0);
                            shortNameList.Add("E");
                        }
                        else if (pipe.Category.Name == "管")
                        {
                            isPipe.Add(1);
                            string pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                            string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                            shortNameList.Add(shortSystemName);
                        }
                        else if (pipe.Category.Name == "風管")
                        {
                            isPipe.Add(2);
                            string pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                            string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                            //shortNameList.Add(shortSystemName);
                            shortNameList.Add(shortSystemName);
                        }
                    }
                    List<int> newTypeList = isPipe.Distinct().ToList();
                    List<string> newList = shortNameList.Distinct().ToList();
                    //先用類別判斷，在用是否共管判斷，就算共管，如果同系統還是得寫一樣的縮寫名稱
                    if (newTypeList.Count() == 1 && newTypeList.First() == 0)
                    {
                        systemType.Set("E");
                    }
                    else if (newTypeList.Count() >= 1 /*&& newTypeList.First() == true*/)
                    {
                        //因為剛好空調也叫M，才需要特別處理
                        if (newList.Count() == 1 && newList.Contains("M"))
                        {
                            //systemType.Set(newList.First());
                            systemType.Set("A");
                        }
                        else if (newList.Count() == 1 && newList.First() != "M")
                        {
                            systemType.Set(newList.First());
                        }
                        //如果為不同系統共管，則設為M
                        else if (newList.Count() > 1)
                        {
                            systemType.Set("M");
                        }
                    }
                    //else if (newTypeList.Count() > 1)
                    //{
                    //    systemType.Set("M");
                    //}
                }
                updateCast = inst;
            }
            catch
            {
                errorOutput += $"套管干涉檢查失敗，ID為 {elem.Id} 的套管無法順利檢查\n";
                //MessageBox.Show($"穿樑套管干涉檢查失敗，ID為{elem.Id}的套管無法順利檢查!");
            }
            return updateCast;
        }
        public Element updateCastMaterial(Element elem, Element linkedBeam)
        {
            FamilyInstance updateCast = null;
            try
            {
                Parameter materialPara = null;
                bool check = false;

                FamilyInstance tempInst = elem as FamilyInstance;
                FamilyInstance tempBeam = linkedBeam as FamilyInstance;
                string instNameCheck = tempInst.Symbol.LookupParameter("API識別名稱").AsString();
                string tempBeamMaterial = tempBeam.StructuralMaterialType.ToString();
                if (checkPara(elem, "貫穿樑材料"))
                {
                    materialPara = elem.LookupParameter("貫穿樑材料");
                    check = true;
                }
                else if (!check)
                {
                    MessageBox.Show("請檢查元件是否有「貫穿樑材料」參數欄位");
                }
                if (instNameCheck.Contains("開口") && tempBeamMaterial == "Steel")
                {
                    materialPara.Set("ST開口");
                }
                else if (instNameCheck.Contains("套管") && tempBeamMaterial == "Concrete")
                {
                    materialPara.Set("RC套管");
                }
                else
                {
                    materialPara.Set("不符合");
                }
                updateCast = tempInst;
            }
            catch
            {
                errorOutput += $"貫穿樑材料更新失敗，ID為 {elem.Id} 的套管無法順利更新\n";
                //MessageBox.Show($"貫穿樑材料更新失敗，ID為{elem.Id}的套管無法順利更新!");
            }
            return updateCast;
        }
        public List<RevitLinkInstance> getMEPLinkInstance(Document doc)
        {
            List<RevitLinkInstance> linkInstanceList = new List<RevitLinkInstance>();
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            //製作針對管、風管、電管的過濾器
            BuiltInCategory[] builts =
            {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_DuctCurves
            };
            List<ElementFilter> filters = new List<ElementFilter>();
            foreach (BuiltInCategory built in builts)
            {
                ElementCategoryFilter filter = new ElementCategoryFilter(built);
                filters.Add(filter);
            }
            LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
            foreach (RevitLinkInstance linkInst in linkCollector)
            {
                Document linkDoc = linkInst.GetLinkDocument();
                if (linkDoc != null)
                {
                    FilteredElementCollector coll = new FilteredElementCollector(linkDoc).WherePasses(categoryFilter).WhereElementIsNotElementType();
                    if (coll.Count() > 0) linkInstanceList.Add(linkInst);
                }
            }
            return linkInstanceList;
        }
    }
}
