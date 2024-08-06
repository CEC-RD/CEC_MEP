#region Namespaces
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;
using System.Linq;
using System;
using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace AutoHangerCreation_ButtonCreate
{
    static class method
    {
        public static IList<Solid> GetTargetSolids(Element element)
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
        public static bool checkPara(Element elem, string paraName)
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
        public static Solid singleSolidFromElement(Element inputElement)
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
        public static RevitLinkInstance getTargetLinkedInstance(Document doc, string linkTilte)
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
        public static List<RevitLinkInstance> findLinkInstWithSlab(Document doc)
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
                        if (linkDoc != null && isLoaded)
                        {
                            FilteredElementCollector linkedSlab = new FilteredElementCollector(linkDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                            if (linkedSlab.Count() == 0) continue;
                            else
                            {
                                //MessageBox.Show($"{linkedInst.Name}中有{linkedSlab.Count()}個樓板");
                                if (!targetLinkedInstances.Contains(linkedInst)) targetLinkedInstances.Add(linkedInst);
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("請檢查外參連結是否載入或有問題!");
            }
            return targetLinkedInstances;
        }
        public static FilteredElementCollector getAllLinkedSlab(Document linkedDoc)
        {
            FilteredElementCollector slabCollector = new FilteredElementCollector(linkedDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            return slabCollector;
        }
        public static void setInstanceParaString(FamilyInstance inst, string paraName)
        {

        }
        public static void updateHangerContent(FamilyInstance inst ,List<Element>pipes)
        {
            try
            {
                List<string> systemName = new List<string>() { "E", "T", "W", "P", "F", "A", "G" };
                List<string> checkName = new List<string>() { "系統別", "系統別" };
                foreach (string st in checkName)
                {
                    if (!checkPara(inst, st)) MessageBox.Show($"請確認{inst.Name}元件中是否缺少{st}參數");
                }
                Parameter systemType = inst.LookupParameter(checkName[0]);
                Parameter pipeNum = inst.LookupParameter(checkName[1]);
                Parameter note = inst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pipes.Count() == 0)
                {
                    systemType.Set("SP");
                    note.Set("SP");
                }
                inst.LookupParameter("干涉管數量").Set(pipes.Count());
                if (pipes.Count() == 1)
                {
                    if (pipes.First().Category.Name == "電管" || pipes.First().Category.Name == "電纜架")
                    {
                        systemType.Set("E");
                        note.Set("E");
                    }
                    else
                    {
                        Element targetPipe = pipes.First();
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
                            note.Set("A");
                        }
                        else
                        {
                            systemType.Set("SP");
                            note.Set("SP");
                        }
                    }
                }
                //如果有共管的狀況
                else if (pipes.Count() >= 2)
                {
                    List<int> isPipe = new List<int>();
                    List<string> shortNameList = new List<string>();
                    foreach (Element pipe in pipes)
                    {
                        if (pipe.Category.Name == "電管" || pipe.Category.Name == "電纜架")
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
                        note.Set("E");
                    }
                    else if (newTypeList.Count() >= 1 /*&& newTypeList.First() == true*/)
                    {
                        //因為剛好空調也叫M，才需要特別處理
                        if (newList.Count() == 1 && newList.Contains("M"))
                        {
                            //systemType.Set(newList.First());
                            systemType.Set("A");
                            note.Set("A");
                        }
                        else if (newList.Count() == 1 && newList.First() != "M")
                        {
                            systemType.Set(newList.First());
                        }
                        //如果為不同系統共管，則設為M
                        else if (newList.Count() > 1)
                        {
                            systemType.Set("M");
                            note.Set("M");
                        }
                    }
                }
            }
            catch
            {
                //errorOutput += $"套管干涉檢查失敗，ID為 {inst.Id} 的套管無法順利檢查\n";
                MessageBox.Show("$套管干涉檢查失敗，ID為 {inst.Id} 的套管無法順利檢查");
            }
        }
        public static Dictionary<ElementId, List<Element>> getHangerSlabDict(Document doc, FamilyInstance inst)
        {
            //其實應該再試試另一個方法，是以連結檔中樓板為主進行變換的方法
            Dictionary<ElementId, List<Element>> hangerSlabDict = new Dictionary<ElementId, List<Element>>();
            try
            {
                List<RevitLinkInstance> slabLinkInstances = findLinkInstWithSlab(doc);
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
                            FilteredElementCollector collectorSB = getAllLinkedSlab(linkedInst.GetLinkDocument());

                            ////針對吊架進行座標變換
                            Transform newTrans = Transform.Identity;
                            BoundingBoxXYZ instBounding = inst.get_BoundingBox(null);
                            Outline outLine = new Outline(inverseTransform.OfPoint(instBounding.Min), inverseTransform.OfPoint(instBounding.Max));
                            //快慢速過濾器結合->和樓板進行碰撞
                            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                            collectorSB.WherePasses(boundingBoxIntersectsFilter);
                            List<Element> tempList = collectorSB.ToList();
                            if (tempList.Count > 0)
                            {
                                hangerSlabDict.Add(inst.Id, tempList);
                            }
                            else
                            {
                                MessageBox.Show("請確認吊架上方是否有樓板可供支撐!");
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("請確認模型中是否有載入對映的土建模型!!");
                }
            }
            catch
            {
                MessageBox.Show("無法判斷吊架與樓板的關係!");
            }
            return hangerSlabDict;
        }
    }
}
