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
    class CopyPartCast : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //針對樓層複製穿梁套管
            Counter.count += 1;
            string messageOut = "";
            int copyNum = 0;
            try
            {
                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Autodesk.Revit.DB.View activeView = doc.ActiveView;
                if (activeView.ViewType != ViewType.FloorPlan)
                {
                    MessageBox.Show("請於平面試圖中使用此功能");
                    return Result.Failed;
                }

                //針對穿梁套管的特性選擇視圖的上一層為複製依據
                Level viewLevel = activeView.GenLevel;
                FilteredElementCollector levelCollector = new FilteredElementCollector(doc).OfClass(typeof(Level));
                List<Element> level_List = levelCollector.OrderBy(x => sortLevelbyHeight(x)).ToList();
                List<string> levelNames = new List<string>(); //用名字來確認篩選排序
                for (int i = 0; i < level_List.Count(); i++)
                {
                    Level le = level_List[i] as Level;
                    levelNames.Add(le.Name);
                }
                int orginIndex = levelNames.IndexOf(viewLevel.Name);
                viewLevel = level_List[orginIndex + 1] as Level;
                Dictionary<Transform, List<FamilyInstance>> linkCastDict = getLinkedCastDictByLevel(doc, viewLevel);

                if (linkCastDict.Keys.Count() == 0)
                {
                    MessageBox.Show("沒有實做的連結模型!");
                }
                //先蒐集在這個層樓中所有穿樑套管的座標點
                List<FamilyInstance> targetInstList = findTargetElementsByLevel(doc, viewLevel);
                List<bool> alreadyExist = new List<bool>();

                using (Transaction trans = new Transaction(doc))
                {
                    int deleteNum = 0;
                    trans.Start($"複製{viewLevel.Name}中的穿梁套管");
                    foreach (FamilyInstance cast in targetInstList)
                    {
                        if (!InstanceExistinLinkByLevel(doc, cast, viewLevel))
                        {
                            doc.Delete(cast.Id);
                            deleteNum += 1;
                        }
                    }
                    MessageBox.Show($"本模型共刪除{viewLevel.Name} 中 {deleteNum} 個多餘的套管");
                    foreach (Transform t in linkCastDict.Keys)
                    {
                        //因為是以外參實體為單位去做實體蒐集
                        List<FamilyInstance> famList = linkCastDict[t];
                        List<ElementId> tempList = new List<ElementId>();
                        Document linkDoc = famList.First().Document;
                        foreach (FamilyInstance inst in famList)
                        {
                            //去除重複者的檢核機制
                            LocationPoint linkedInstLocate = inst.Location as LocationPoint;
                            XYZ linkedPt = linkedInstLocate.Point;
                            linkedPt = TransformPoint(linkedPt, t);
                            //準備要複製的這個跟已經有的去做比對，如果已經在模型中，則不複製，不在則需要複製
                            if (!InstanceExistByLevel(doc, linkedPt,viewLevel))
                            {
                                tempList.Add(inst.Id);
                            }
                        }
                        ICollection<ElementId> temp = tempList;
                        copyNum += temp.Count();
                        //有少才複製與顯示訊息
                        if (temp.Count > 0)
                        {
                            messageOut += $"來自 {linkDoc.Title} 的外參中少了 {tempList.Count()} 個套管\n";
                            //MessageBox.Show($"{linkDoc.Title}的外參中少了{tempList.Count()}個套管");
                            CopyPasteOptions options = new CopyPasteOptions();
                            ElementTransformUtils.CopyElements(linkDoc, temp, doc, t, options);
                        }

                    }
                    messageOut += $"連結模型中位於{viewLevel.Name}的穿樑套管已全數複製完成，共複製了 {copyNum} 個套管!\n";
                    trans.Commit();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }

            MessageBox.Show(messageOut);
            return Result.Succeeded;
        }
        public Level getTargetLevelfromLink(RevitLinkInstance linkInst, Level localLevel)
        {
            Level targetLevel = null;
            Document linkDoc = linkInst.GetLinkDocument();
            FilteredElementCollector levelCollector = new FilteredElementCollector(linkDoc).OfClass(typeof(Level));
            
            foreach (Element e in levelCollector)
            {
                Element localElement = localLevel as Element;
                Level tempLevel = e as Level;
                if (e.Name == localLevel.Name || tempLevel.ProjectElevation == localLevel.ProjectElevation)
                {
                    targetLevel = tempLevel;
                    break;
                }
            }
            if (targetLevel == null) MessageBox.Show("請確認外參模型與本機端模型的高程與命名原則是否一致");
            return targetLevel;
        }
        public Dictionary<Transform, List<FamilyInstance>> getLinkedCastDictByLevel(Document doc, Level localLevel)
        {
            Dictionary<Transform, List<FamilyInstance>> linkCastBeamDict = new Dictionary<Transform, List<FamilyInstance>>();
            List<RevitLinkInstance> linkedMEP_files = getMEPLinkInstance(doc);
            foreach (RevitLinkInstance linkInst in linkedMEP_files)
            {
                List<FamilyInstance> targetList = new List<FamilyInstance>();
                Transform linkTrans = linkInst.GetTotalTransform();
                Document linkDoc = linkInst.GetLinkDocument();

                Level linkedLevel = getTargetLevelfromLink(linkInst, localLevel);
                ElementLevelFilter levelFilter = new ElementLevelFilter(linkedLevel.Id);
                FilteredElementCollector linkCastCollector = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(Instance)).WherePasses(levelFilter);
                foreach (FamilyInstance inst in linkCastCollector)
                {
                    //針對checkName一定要確認是否為null，因為有些元件沒有此參數
                    Parameter checkName = inst.Symbol.LookupParameter("API識別名稱");
                    if (checkName != null && checkName.AsString().Contains("CEC-穿樑"))
                    {
                        targetList.Add(inst);
                    }
                }
                //MessageBox.Show($"{linkDoc.Title}的{localLevel.Name}中共有{ targetList.Count()}個套管");
                //如果有蒐集到套管
                if (targetList.Count() > 0)
                {
                    if (!linkCastBeamDict.Keys.Contains(linkTrans))
                    {
                        linkCastBeamDict.Add(linkTrans, targetList);
                    }
                    else if (linkCastBeamDict.Keys.Contains(linkTrans))
                    {
                        //continue;
                        foreach (FamilyInstance inst in targetList)
                        {
                            linkCastBeamDict[linkTrans].Add(inst);
                        }
                    }
                }
            }
            return linkCastBeamDict;
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
        public List<FamilyInstance> findTargetElementsByLevel(Document doc, Level le)
        {
            //RC套管跟SC開口的內部名稱是不同的
            string internalNameST = "CEC-穿樑開口";
            string internalNameRC = "CEC-穿樑套管";
            List<FamilyInstance> castInstances = new List<FamilyInstance>();
            try
            {
                FilteredElementCollector coll = new FilteredElementCollector(doc);
                ElementCategoryFilter castCate_Filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementClassFilter castInst_Filter = new ElementClassFilter(typeof(Instance));
                LogicalAndFilter andFilter = new LogicalAndFilter(castCate_Filter, castInst_Filter);
                ElementLevelFilter levelFilter = new ElementLevelFilter(le.Id);
                coll.WherePasses(andFilter).WherePasses(levelFilter).WhereElementIsNotElementType().ToElements(); //找出模型中實做的穿樑套管元件
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
                MessageBox.Show("蒐集套管發生問題!");
            }
            return castInstances;
        }
        public List<FamilyInstance> findLinkedCast(Document doc)
        {
            //原本想用level找，但此方法不可行，因為外參中的level已為不同元件
            List<RevitLinkInstance> linkedMEP_files = getMEPLinkInstance(doc);
            List<FamilyInstance> targetList = new List<FamilyInstance>();
            //ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);
            //針對所有在此檔案中的機電外參檔找尋套管
            foreach (RevitLinkInstance linkInst in linkedMEP_files)
            {
                Transform linkTrans = linkInst.GetTotalTransform();
                Document linkDoc = linkInst.GetLinkDocument();
                FilteredElementCollector linkCastCollector = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(Instance));
                foreach (FamilyInstance inst in linkCastCollector)
                {
                    //針對checkName一定要確認是否為null，因為有些元件沒有此參數
                    Parameter checkName = inst.Symbol.LookupParameter("API識別名稱");
                    if (checkName != null && checkName.AsString().Contains("CEC-穿樑"))
                    {
                        targetList.Add(inst);
                    }
                }
                //如果有蒐集到套管，則將套管ID
            }
            return targetList;
        }
        public static XYZ TransformPoint(XYZ point, Transform transform)
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
        public bool InstanceExistByLevel(Document document, XYZ point,Level localLevel)
        {
            ElementLevelFilter levelFilter = new ElementLevelFilter(localLevel.Id);
            FilteredElementCollector filteredElementCollector = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PipeAccessory).WherePasses(levelFilter);
            List<FamilyInstance> targetList = new List<FamilyInstance>();
            foreach (FamilyInstance inst in filteredElementCollector)
            {
                Parameter p = inst.Symbol.LookupParameter("API識別名稱");
                if (p != null && p.AsString().Contains("CEC-穿樑"))
                {
                    targetList.Add(inst);
                }
            }
            //foreach (Instance instance in filteredElementCollector)
            foreach (FamilyInstance instance in targetList)
            {
                Element element = instance as Element;
                Location loc = instance.Location;
                LocationPoint locpoint = loc as LocationPoint;
                if (locpoint.Point.IsAlmostEqualTo(point, 0.001))
                {
                    return true;
                }
            }
            return false;
        }
        public bool InstanceExistinLinkByLevel(Document document, FamilyInstance cast, Level localLevel)
        {
            //用既有的套管和外參中的套管檢查，如果名稱(Symbol)一樣，位置也一樣，則刪除
            //蒐集本機端的套管資訊
            string castSymbolName = cast.Symbol.Name;
            LocationPoint castLocate = cast.Location as LocationPoint;
            XYZ castPt = castLocate.Point;

            Dictionary<Transform, List<FamilyInstance>> linkCastDict = getLinkedCastDictByLevel(document, localLevel);
            List<XYZ> linkCastLocateList = new List<XYZ>();
            foreach (Transform trans in linkCastDict.Keys)
            {
                List<FamilyInstance> instListLinked = linkCastDict[trans];
                foreach (FamilyInstance inst in instListLinked)
                {
                    LocationPoint instLocate = inst.Location as LocationPoint;
                    XYZ instPt = instLocate.Point;
                    instPt = TransformPoint(instPt, trans);
                    linkCastLocateList.Add(instPt);
                    string instSymbolName = inst.Symbol.Name;
                    if (instPt.IsAlmostEqualTo(castPt) && castSymbolName == instSymbolName)
                    {
                        //true 表示在外參中有，但在實體檔中也有
                        return true;
                    }
                }
            }
            //false 表示在外參中沒有，但在實體檔中有，需要被刪掉
            return false;
        }
        public double sortLevelbyHeight(Element element)
        {
            Level tempLevel = element as Level;
            double levelHeight = tempLevel.Elevation;
            return levelHeight;
        }
    }
}
