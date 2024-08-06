#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using System.Threading;
#endregion

namespace CEC_NumRule
{
    class method
    {
        public BuiltInCategory[] builts =
            {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_CableTray
            };
        public string paraName = "系統別_映射";
        private LogicalOrFilter createOrFilter(BuiltInCategory[] builts)
        {
            List<ElementFilter> filters = new List<ElementFilter>();
            foreach (BuiltInCategory built in builts)
            {
                ElementCategoryFilter filter = new ElementCategoryFilter(built);
                filters.Add(filter);
            }
            LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
            return categoryFilter;

        }
        public List<customSystem> GetCustomSystems(Document doc, BuiltInCategory built)
        {
            List<customSystem> targetList = new List<customSystem>();
            List<Element> targetCollector = new FilteredElementCollector(doc).OfCategory(built).WhereElementIsElementType().OrderBy(x => x.Name).ToList();
            foreach (Element e in targetCollector)
            {
                customSystem tempSystem = new customSystem()
                {
                    systemType = e,
                    systemName = e.Name,
                    //targetSystemName = ""
                };
                IList<Guid> idList = e.GetEntitySchemaGuids();
                Schema schema = openingRuleScheme.getCivilSystemSchema();
                Entity entity = e.GetEntity(schema);
                if (entity.IsValid())//有找到綁定的數據實例才使用
                {
                    tempSystem.targetSystemName = openingRuleScheme.retrieveDataFromElement(schema, e);
                }
                else
                {
                    tempSystem.targetSystemName = "";
                }
                targetList.Add(tempSystem);
            }
            return targetList;
        }
        //蒐集套管=>利用"穿牆"、"穿梁"、"穿版"來蒐集不同部分的開口
        public List<Element> getTargetElement(Document doc, List<string> triggers)
        {
            //依照trigger蒐集相對應的套管
            List<Element> targetList = new List<Element>();
            foreach (string st in triggers)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType();
                foreach (Element e in collector)
                {
                    FamilySymbol symbol = doc.GetElement(e.GetTypeId()) as FamilySymbol;
                    Parameter para = symbol.LookupParameter("API識別名稱");
                    if (para != null && para.AsString().Contains(st))
                    {
                        targetList.Add(e);
                    }
                }
            }
            return targetList;
        }
        public List<Element> getTargetElement(Document doc, List<string> triggers, Autodesk.Revit.DB.View activeView)
        {
            //依照trigger蒐集相對應的套管
            List<Element> targetList = new List<Element>();
            if (activeView.ViewType != ViewType.FloorPlan) MessageBox.Show("請在平面視圖使用此功能");
            ViewPlan viewPlan = activeView as ViewPlan;
            Level genLevel = viewPlan.GenLevel;
            ElementLevelFilter levelFilter = new ElementLevelFilter(genLevel.Id);

            foreach (string st in triggers)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(FamilyInstance)).WherePasses(levelFilter).WhereElementIsNotElementType();
                foreach (Element e in collector)
                {
                    FamilySymbol symbol = doc.GetElement(e.GetTypeId()) as FamilySymbol;
                    Parameter para = symbol.LookupParameter("API識別名稱");
                    if (para != null && para.AsString().Contains(st))
                    {
                        targetList.Add(e);
                    }
                }
            }
            return targetList;
        }
        public List<Element> getTargetElement(Document doc, string trigger, Autodesk.Revit.DB.View activeView)
        {
            List<Element> targetList = new List<Element>();
            if (activeView.ViewType != ViewType.FloorPlan) MessageBox.Show("請在平面視圖使用此功能");
            ViewPlan viewPlan = activeView as ViewPlan;
            Level genLevel = viewPlan.GenLevel;
            ElementLevelFilter levelFilter = new ElementLevelFilter(genLevel.Id);

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(FamilyInstance)).WherePasses(levelFilter).WhereElementIsNotElementType();
            foreach (Element e in collector)
            {
                FamilySymbol symbol = doc.GetElement(e.GetTypeId()) as FamilySymbol;
                Parameter para = symbol.LookupParameter("API識別名稱");
                if (para != null && para.AsString().Contains(trigger))
                {
                    targetList.Add(e);
                }
            }
            return targetList;
        }
        private List<RevitLinkInstance> getMEPLinkInstance(Document doc)
        {
            List<RevitLinkInstance> linkInstanceList = new List<RevitLinkInstance>();
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            LogicalOrFilter categoryFilter = createOrFilter(builts);
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

        public Dictionary<ElementId, List<Element>> getCastDict(Document doc, List<Element> castList)
        {
            Dictionary<ElementId, List<Element>> targetDict = new Dictionary<ElementId, List<Element>>();
            List<RevitLinkInstance> MEPlinkInstes = getMEPLinkInstance(doc);
            //先蒐集本機管
            using (ProgressUI progerssUI = new ProgressUI(castList.Count()))
            {
                //蒐集外參管
                foreach (Element e in castList)
                {
                    List<Element> targetList = collectPipesByCast(doc, e);
                    if (targetList.Count() != 0)
                    {
                        targetDict.Add(e.Id, targetList);
                    }
                    foreach (RevitLinkInstance linkInst in MEPlinkInstes)
                    {
                        List<Element> linktList = collectPipesByCast(linkInst, e);
                        if (targetDict.Keys.Contains(e.Id) && linktList.Count() != 0)
                        {
                            targetDict[e.Id].AddRange(linktList);
                        }
                        else if (!targetDict.Keys.Contains(e.Id) && linktList.Count() != 0)
                        {
                            targetDict.Add(e.Id, linktList);
                        }
                    }
                    if (progerssUI.Update()) break;
                }
            }

            return targetDict;
        }
        private List<Element> collectPipesByCast(Document doc, Element cast)
        {
            List<Element> targetList = new List<Element>();
            LogicalOrFilter cateFilter = createOrFilter(builts);
            FilteredElementCollector coll = new FilteredElementCollector(doc)/*.OfClass(typeof(Instance))*/.WherePasses(cateFilter).WhereElementIsNotElementType();
            LogicalAndFilter castFilter = createCastFilter(cast);
            coll = coll.WherePasses(castFilter);
            if (coll.Count() != 0)
            {
                targetList = coll.ToList();
            }
            return targetList;
        }
        private List<Element> collectPipesByCast(RevitLinkInstance linkInst, Element cast)
        {
            Document linkDoc = linkInst.GetLinkDocument();
            Transform linkTrans = linkInst.GetTotalTransform();
            List<Element> targetList = new List<Element>();
            LogicalOrFilter cateFilter = createOrFilter(builts);
            LogicalAndFilter andFilter = createCastFilter(cast, linkTrans);
            FilteredElementCollector coll = new FilteredElementCollector(linkDoc)/*.OfClass(typeof(Instance))*/.WherePasses(cateFilter).WherePasses(andFilter).WhereElementIsNotElementType();
            if (coll.Count() != 0)
            {
                targetList = coll.ToList();
            }
            return targetList;
        }
        //public List<Element> pipesConflict(Element e, FilteredElementCollector )
        //{

        //}
        private IList<Solid> GetTargetSolids(Element element)
        {
            List<Solid> solids = new List<Solid>();
            Options options = new Options();
            //預設為不包含不可見元件，因此改成true
            options.ComputeReferences = true;
            options.DetailLevel = ViewDetailLevel.Fine;
            options.IncludeNonVisibleObjects = false;
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
        private Solid singleSolidFromElement(Element inputElement)
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
            string errorID = "";
            try
            {
                if (solidCount == 1)
                {
                    errorID = "1";
                    solidResult = fromElement[0];
                }
                else if (solidCount > 1)
                {
                    errorID = "2";
                    solidResult =
                        BooleanOperationsUtils.ExecuteBooleanOperation(fromElement[0], fromElement[1], BooleanOperationsType.Union);
                }

                if (solidCount > 2)
                {
                    errorID = "3";
                    for (int i = 2; i < solidCount; i++)
                    {
                        solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(solidResult, fromElement[i], BooleanOperationsType.Union);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"來自{inputElement.Document.Title}的{inputElement.Name}:{inputElement.Id}在幾何轉換上出問題，請檢查{errorID}");
            }

            return solidResult;
        }
        public Solid singleSolidFromWall(Element element)
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
        private LogicalAndFilter createCastFilter(Element cast)
        {
            BoundingBoxXYZ castBounding = cast.get_BoundingBox(null);
            Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);

            Solid castSolid = singleSolidFromElement(cast);
            ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
            LogicalAndFilter andFilter = new LogicalAndFilter(boundingBoxIntersectsFilter, solidFilter);
            return andFilter;
        }
        private LogicalAndFilter createCastFilter(Element cast, Transform trans)
        {
            Solid castSolid = singleSolidFromElement(cast);
            castSolid = SolidUtils.CreateTransformed(castSolid, trans.Inverse);
            BoundingBoxXYZ solidBounding = castSolid.GetBoundingBox();
            XYZ solidCenter = castSolid.ComputeCentroid();
            Transform transform = Transform.Identity;
            transform.Origin = solidCenter;
            Outline outLine = new Outline(transform.OfPoint(solidBounding.Min), transform.OfPoint(solidBounding.Max));
            BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
            ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
            LogicalAndFilter andFilter = new LogicalAndFilter(boxIntersectsFilter, solidFilter);
            return andFilter;
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
        public void reUpdateCastSystem(Dictionary<ElementId, List<Element>> castDict, Document doc)
        {
            Schema schema = Schema.Lookup(openingRuleScheme.civilSystemSchemaGUID);
            using (ProgressUI progerssUI = new ProgressUI(castDict.Keys.Count()))
            {
                foreach (ElementId id in castDict.Keys)
                {
                    string errorOut = "";
                    List<string> systemList = new List<string>();
                    foreach (Element e in castDict[id])
                    {
                        Element ele = getEntityElement(e);
                        Entity entity = ele.GetEntity(schema);

                        if (entity.IsValid())
                        {
                            string targetName = openingRuleScheme.retrieveDataFromElement(schema, ele);
                            if (!systemList.Contains(targetName))
                            {
                                systemList.Add(targetName);
                            }
                        }
                    }
                    FilteredElementCollector coll = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType();
                    string paraName = "系統別_映射";
                    string paraCheck = "不更新系統別";
                    Element castToSet = doc.GetElement(id);
                    //MessageBox.Show(errorOut);
                    if (!checkPara(castToSet, paraCheck))
                    {
                        FamilyInstance instance = castToSet as FamilyInstance;
                        FamilySymbol symbol = instance.Symbol;
                        MessageBox.Show($"請確認 {symbol.Name} 中使否存在 {paraCheck} 參數");
                        break;
                    }
                    if (!checkPara(castToSet, paraName))
                    {
                        FamilyInstance instance = castToSet as FamilyInstance;
                        FamilySymbol symbol = instance.Symbol;
                        MessageBox.Show($"請確認 {symbol.Name} 中使否存在 {paraName} 參數");
                        break;
                    }
                    else if (castToSet.LookupParameter(paraCheck).AsInteger() == 0)
                    {
                        Parameter para = castToSet.LookupParameter(paraName);
                        if (systemList.Count == 1)
                        {
                            para.Set(systemList.First());
                        }
                        else if (systemList.Count > 1)
                        {
                            string result = "";
                            foreach (string st in systemList)
                            {
                                result += $"{st},";
                            }
                            char[] charsToTrim = { ',', '.', ' ' };
                            result = result.TrimEnd(charsToTrim);
                            para.Set(result);
                        }
                        else
                        {
                            para.Set("未知");
                        }
                    }
                    if (progerssUI.Update()) break;
                }
            }
        }

        public void reUpdateCastSystem(ElementId id, List<Element> elements, Document doc)
        {
            Schema schema = Schema.Lookup(openingRuleScheme.civilSystemSchemaGUID);
            string errorOut = "";
            List<string> systemList = new List<string>();
            foreach (Element e in elements)
            {
                Element ele = getEntityElement(e);
                Entity entity = ele.GetEntity(schema);
                if (entity.IsValid())
                {
                    string targetName = openingRuleScheme.retrieveDataFromElement(schema, ele);
                    if (!systemList.Contains(targetName))
                    {
                        systemList.Add(targetName);
                    }
                }
            }
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType();
            string paraName = "系統別_映射";
            string paraCheck = "不更新系統別";
            Element castToSet = doc.GetElement(id);
            //MessageBox.Show(errorOut);
            //if (!checkPara(castToSet, paraCheck))
            //{
            //    FamilyInstance instance = castToSet as FamilyInstance;
            //    FamilySymbol symbol = instance.Symbol;
            //    MessageBox.Show($"請確認 {symbol.Name} 中使否存在 {paraCheck} 參數");
            //    break;
            //}
            //if (!checkPara(castToSet, paraName))
            //{
            //    FamilyInstance instance = castToSet as FamilyInstance;
            //    FamilySymbol symbol = instance.Symbol;
            //    MessageBox.Show($"請確認 {symbol.Name} 中使否存在 {paraName} 參數");
            //    break;
            //}
            if (castToSet.LookupParameter(paraCheck).AsInteger() == 0)
            {
                Parameter para = castToSet.LookupParameter(paraName);
                if (systemList.Count == 1)
                {
                    para.Set(systemList.First());
                }
                else
                {
                    ////para.Set(systemList.First());
                    //para.Set(systemList.Count.ToString());
                    para.Set("未知");
                }
            }
        }

        public Element getEntityElement(Element element)
        {
            Category category = element.Category;
            Element targetElement = null;
            Document doc = element.Document;
            //MessageBox.Show(doc.Title);
            //element.LookupParameter("標註").Set(doc.Title);
            Category pipe = Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(doc, BuiltInCategory.OST_Conduit);
            Category cableTray = Category.GetCategory(doc, BuiltInCategory.OST_CableTray);
            if (category.Name == pipe.Name || category.Name == duct.Name)
            {
                MEPCurve mepCurve = element as MEPCurve;
                MEPSystem mepSystem = mepCurve.MEPSystem;
                ElementId id = mepSystem.GetTypeId();
                targetElement = doc.GetElement(id);
            }
            else if (category.Name == conduit.Name || category.Name == cableTray.Name)
            {
                ElementId id = element.GetTypeId();
                targetElement = doc.GetElement(id);
            }
            return targetElement;
        }

        private List<Element> findAllLevel(Document doc)
        {
            List<Element> targetList = null;
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
            levelCollector.OfClass(typeof(Level))/*.WherePasses(level_Filter)*/.WhereElementIsNotElementType();
            //MessageBox.Show(levelCollector.Count().ToString());
            //foreach (Element e in levelCollector)
            //{
            //    MessageBox.Show($"{doc.Title}+{e.Name}");
            //    Level level = e as Level;
            //    if (level != null)
            //    {
            //        targetList.Add(level);
            //    }
            //}
            //MessageBox.Show(targetList.Count.ToString());
            //if (targetList == null) MessageBox.Show("...");
            //MessageBox.Show(targetList.Count().ToString());
            targetList = levelCollector.ToList();
            return targetList;
        }
        public List<linkObject> getLinkObjects(Document doc, BuiltInCategory builts, Autodesk.Revit.DB.View activeView)
        {
            string outPut = "";
            List<linkObject> targetLinkObject = new List<linkObject>();
            ViewPlan viewPlan = activeView as ViewPlan;
            Level genLevel = viewPlan.GenLevel;
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            foreach (RevitLinkInstance linkInst in linkCollector)
            {
                Level tempLevel = null;
                Document linkDoc = linkInst.GetLinkDocument();
                Transform transform = linkInst.GetTotalTransform();
                List<Element> levelList = findAllLevel(linkDoc);

                if (levelList != null || levelList.Count != 0)
                {
                    foreach (Element element in levelList)
                    {
                        outPut += $"我是來自{linkDoc.Title}專案的「{element.Name}」樓層";
                        Level level = element as Level;
                        //if (level == null) MessageBox.Show($"來自「{linkDoc.Title}」檔案中的{element.Name}無法轉型成樓層");
                        if (level.Name == genLevel.Name || level.ProjectElevation == genLevel.ProjectElevation || level.Elevation==genLevel.Elevation)
                        {
                            tempLevel = level;
                            break;
                        }
                    }
                }
                if (tempLevel == null) MessageBox.Show("請確認視圖參考樓層的「名稱」與「高程」是否與外參模型一致");

                //MessageBox.Show(tempLevel.Name);
                ElementLevelFilter levelFilter = new ElementLevelFilter(tempLevel.Id);
                if (linkDoc != null)
                {
                    //2024.01.26修正，因牆底參考樓層有可能不一樣，避免用同一樓層作為干涉依據
                    FilteredElementCollector coll = new FilteredElementCollector(linkDoc).OfCategory(builts)/*.WherePasses(levelFilter)*/.WhereElementIsNotElementType();
                    if (coll.Count() != 0)
                    {
                        foreach (Element e in coll)
                        {
                            linkObject newObject = new linkObject(linkDoc, e.Id, e, transform);
                            try
                            {
                                newObject.centerPt = getCenterPoint(newObject);
                            }
                            catch
                            {
                                MessageBox.Show($"{e.Category.Name}中的{e.Name}:{e.Id}無法取得中心點");
                            }
                            //if (newObject.centerPt == null)
                            //{
                            //    MessageBox.Show($"來自 {linkDoc.Title} 的 {newObject.linkID} 無法取得物件中心點");
                            //    continue;
                            //}
                            if (newObject.centerPt != null)
                            {
                                targetLinkObject.Add(newObject);
                            }
                        }
                    }
                }
            }
            MessageBox.Show("測試");
            targetLinkObject = targetLinkObject.OrderBy(x => x.centerPt.X).ThenBy(x => x.centerPt.Y).ToList();
            return targetLinkObject;
        }
        private XYZ getCenterPoint(linkObject linkobject)
        {
            //先設定版、牆、樑品類(不同元件中心點有不同的算法)
            XYZ centerPoint = null;
            Transform trans = linkobject.transform;
            Category slab = Category.GetCategory(linkobject.linkDoc, BuiltInCategory.OST_Floors);
            Category wall = Category.GetCategory(linkobject.linkDoc, BuiltInCategory.OST_Walls);
            Category beam = Category.GetCategory(linkobject.linkDoc, BuiltInCategory.OST_StructuralFraming);
            if (linkobject.linkElement.Category.Id == wall.Id || linkobject.linkElement.Category.Id == beam.Id)
            {
                LocationCurve locationCurve = linkobject.linkElement.Location as LocationCurve;
                if (locationCurve != null)
                {
                    Curve curve = locationCurve.Curve;
                    curve = curve.CreateTransformed(trans);
                    centerPoint = curve.Evaluate(0.5, true);
                }
            }
            else if (linkobject.linkElement.Category.Id == slab.Id)
            {
                Solid slabSolid = singleSolidFromElement(linkobject.linkElement);
                slabSolid = SolidUtils.CreateTransformed(slabSolid, trans);
                centerPoint = slabSolid.ComputeCentroid();
            }
            return centerPoint;
        }
        public List<Element> getCastInLinkObject(Document doc, linkObject linkObject, string trigger)
        {
            //針對本機端的套管去做編號更新，繼續編號的原則不使用補位的方式
            //step1. 針對版、牆、樑去蒐集元素
            Transform linkTrans = linkObject.transform;
            Element element = linkObject.linkElement;
            BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);
            Solid elementSolid = singleSolidFromElement(element);
            List<Element> targetList = new List<Element>();
            if (elementSolid != null)
            {
                elementSolid = SolidUtils.CreateTransformed(elementSolid, linkTrans);
                XYZ solidCenter = elementSolid.ComputeCentroid();
                ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(elementSolid);
                Transform newTrans = Transform.Identity;
                newTrans.Origin = solidCenter;
                XYZ targetMaxPt = TransformPoint(boundingBox.Max, linkTrans);
                XYZ targetMinPt = TransformPoint(boundingBox.Min, linkTrans);

                //Outline outLine = new Outline(newTrans.OfPoint(boundingBox.Min), newTrans.OfPoint(boundingBox.Max));
                Outline outLine = new Outline(targetMinPt, targetMaxPt);
                //createSolidFromBBox(doc, targetMaxPt, targetMinPt);
                //createSolidFromBBox(doc, newTrans.OfPoint(boundingBox.Max), newTrans.OfPoint(boundingBox.Min));
                BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                Autodesk.Revit.DB.View activeView = doc.ActiveView;
                ViewPlan activePlan = null;
                if (activeView.ViewType == ViewType.FloorPlan)
                {
                    activePlan = activeView as ViewPlan;
                }
                else
                {
                    MessageBox.Show("請在平面視圖使用此功能");
                }
                Level referLevel = activePlan.GenLevel;
                ElementLevelFilter levelFilter = new ElementLevelFilter(referLevel.Id);
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(FamilyInstance)).WherePasses(levelFilter).WherePasses(boundingBoxIntersectsFilter).WherePasses(solidFilter);


                foreach (Element e in collector)
                {
                    FamilyInstance inst = e as FamilyInstance;
                    FamilySymbol symbol = inst.Symbol;
                    Parameter para = symbol.LookupParameter("API識別名稱");
                    if (para != null && para.AsString().Contains(trigger))
                    {
                        targetList.Add(e);
                    }
                }
                targetList = targetList/*.OrderBy(x=>x.LookupParameter(paraName).AsString())*/.OrderBy(x => getLocationPt(x).X).ThenBy(x => getLocationPt(x).Y).ToList();
            }
            return targetList;
        }
        private XYZ TransformPoint(XYZ point, Transform transform)
        {
            double x = point.X;
            double y = point.Y;
            double z = point.Z;

            //transform basis of the old coordinate system in the new coordinate // system
            XYZ b0 = transform.get_Basis(0);
            XYZ b1 = transform.get_Basis(1);
            XYZ b2 = transform.get_Basis(2);
            XYZ origin = transform.Origin;

            //transform the origin of the old coordinate system in the new 
            //coordinate system
            double xTemp = x * b0.X + y * b1.X + z * b2.X + origin.X;
            double yTemp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y;
            double zTemp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z;

            return new XYZ(xTemp, yTemp, zTemp);
        }
        private XYZ getLocationPt(Element e)
        {
            LocationPoint locationPoint = e.Location as LocationPoint;
            XYZ targetPoint = null;
            if (locationPoint != null)
            {
                targetPoint = locationPoint.Point;
            }
            return targetPoint;
        }
        private void createSolidFromBBox(Document doc, XYZ maxPt, XYZ minPt)
        {
            XYZ pt0 = new XYZ(minPt.X, minPt.Y, minPt.Z);
            XYZ pt1 = new XYZ(maxPt.X, minPt.Y, minPt.Z);
            XYZ pt2 = new XYZ(maxPt.X, maxPt.Y, minPt.Z);
            XYZ pt3 = new XYZ(minPt.X, maxPt.Y, minPt.Z);
            List<CurveLoop> loopList = new List<CurveLoop>();
            Line edge0 = Line.CreateBound(pt0, pt1);
            Line edge1 = Line.CreateBound(pt1, pt2);
            Line edge2 = Line.CreateBound(pt2, pt3);
            Line edge3 = Line.CreateBound(pt3, pt0);
            List<Curve> edges = new List<Curve>();
            edges.Add(edge0);
            edges.Add(edge1);
            edges.Add(edge2);
            edges.Add(edge3);
            CurveLoop baseLoop = CurveLoop.Create(edges);
            loopList.Add(baseLoop);
            double height = maxPt.Z - minPt.Z;
            Solid SoldBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, height);
            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            ds.ApplicationId = "Test";
            ds.ApplicationDataId = "testBox";
            List<GeometryObject> GeoList = new List<GeometryObject>();
            GeoList.Add(SoldBox); // <-- the solid created for the intersection can be used here
            ds.SetShape(GeoList);
            //ds.SetName("ID_testBox");
        }
    }

    static class openingRuleScheme
    {
        private const string vendorID = "CEC_chliu";
        private const string civilSystemRuleShemaName = "CEC_CivilSystem";
        public static Guid civilSystemSchemaGUID = new Guid("1E82C544-7144-42E9-B858-36CC4375A476");
        private const string fieldName = "civilSystemMap";

        public static Schema getCivilSystemSchema()
        {
            Schema schema = Schema.Lookup(civilSystemSchemaGUID);
            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(civilSystemSchemaGUID);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);
                schemaBuilder.SetVendorId(vendorID);
                schemaBuilder.SetSchemaName(civilSystemRuleShemaName);

                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(fieldName, typeof(string));
                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        public static void setEntityToElement(Document doc, Schema schema, Element e, string value)
        {
            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("寫入Schema資料");
                Entity entity = new Entity(schema);
                Field fieldToSet = schema.GetField(fieldName);
                entity.Set(fieldToSet, value);
                e.SetEntity(entity);
                trans.Commit();
            }
        }

        public static string retrieveDataFromElement(Schema schema, Element e)
        {
            Entity retrieveEntity = e.GetEntity(schema);
            string retrieveData = retrieveEntity.Get<string>(schema.GetField(fieldName));
            return retrieveData;
        }
    }
    public class customSystem
    {
        //管系統類型、風管系統類型、電管類型、電纜架類型
        public Element systemType { get; set; }
        public string systemName { get; set; }
        public string targetSystemName { get; set; }
    }
    public class linkObject
    {
        public linkObject(Document doc, ElementId id, Element element, Transform trans)
        {
            linkDoc = doc;
            linkID = id;
            linkElement = element;
            transform = trans;
        }
        public Document linkDoc { get; set; }
        public ElementId linkID { get; set; }
        public Element linkElement { get; set; }
        public Transform transform { get; set; }
        public XYZ centerPt { get; set; }
    }
}
