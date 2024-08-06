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
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using System.IO;
using System.Threading;
#endregion

namespace CEC_PreFabric
{
    //【程式架構】
    //0.讓這隻程式只能再3D視圖中使用
    //0.先彈出UI視窗
    //1.先檢查3D視圖中的管是否都有裁切料號參數，且不為""，如果有則報錯
    //2.針對裁切料號進行重新編組
    //3.利用選擇的elements來改變ISO的orientation
    [Transaction(TransactionMode.Manual)]
    public class ReNumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            List<ElementId> tagsToDelete = new List<ElementId>();
            bool checkPara = true;
            List<string> paraName1 = new List<string>() { "【預組】系統別", "【預組】樓層", "【預組】區域", "【預組】編號" };
            List<Element> alreadySetList = new List<Element>();
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            if (activeView.ViewType != ViewType.ThreeD)
            {
                MessageBox.Show("視圖並非3D視圖，請於3D視圖中使用");
                return Result.Failed;
            }
            foreach (string st in paraName1)
            {
                BindingMap bm = doc.ParameterBindings;
                DefinitionBindingMapIterator itor = bm.ForwardIterator();
                itor.Reset();
                Definition d = null;
                ElementBinding elemBind = null;
                while (itor.MoveNext())
                {
                    d = itor.Key;
                    if (d.Name == st)
                    {
                        elemBind = (ElementBinding)itor.Current;
                        break;
                    }
                }
                if (d.Name != st)
                {
                    checkPara = false;
                    break;
                }
            }
            if (checkPara == false)
            {
                MessageBox.Show("請檢查專案中是否有匯入【預組】系統別、【預組】樓層、【預組】區域、【預組】編號 等四個共用參數");
            }
            //確定有匯入共用參數後，先彈出UI讓使用者設定參數
            PreFabricUI ui = new PreFabricUI();
            ui.viewNameTextBox.IsReadOnly = true;
            ui.viewTemplateComboBox.IsReadOnly = true;
            ui.regionName.IsReadOnly = true;
            ui.levelName.IsReadOnly = true;
            ui.ShowDialog();
            if (ui.DialogResult == false)
            {
                return Result.Failed;
            }
            int keyToSet = 0;
            try
            {
                keyToSet = Convert.ToInt32(ui.startingNumTextBox.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("輸入的起始值為無效值(不可輸入文字或特殊符號)!");
                return Result.Failed;
            }
            using (TransactionGroup transGroup = new TransactionGroup(doc))
            {
                transGroup.Start("預組管重新編號");
                while (true)
                {
                    try
                    {
                        //確定有匯入共用參數之後，蒐集視圖內所有的管
                        ElementCategoryFilter pipeFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
                        FilteredElementCollector viewElementFilter = new FilteredElementCollector(doc, activeView.Id);
                        List<Element> pipeList = new List<Element>();
                        Category pipeCate = Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves);
                        ISelectionFilter pipeSelectFilter = new PipeSelectionFilter(doc);
                        Reference refer = uidoc.Selection.PickObject(ObjectType.Element, pipeSelectFilter, "請選擇要重新編號的管");
                        Element pickPipe = doc.GetElement(refer.ElementId);
                        if (pickPipe.LookupParameter(paraName1[3]).AsString() != "")
                        {
                            MessageBox.Show("選中的管已有編號，請重新選擇!");
                        }
                        else
                        {
                            foreach (Element e in viewElementFilter)
                            {
                                if (e.Category != null)
                                {
                                    if (e.Category.Id == pipeCate.Id)
                                    {
                                        pipeList.Add(e);
                                    }
                                }
                            }
                            //針對同樣大小的管材
                            List<Element> sameList = new List<Element>(); ;
                            string elemName = pickPipe.Name;
                            string elemSize = getPipeDiameter(pickPipe);
                            double elemLength = Math.Round(pickPipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), 2);
                            foreach (Element ee in pipeList)
                            {
                                double tempLength = Math.Round(ee.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), 2);
                                if (elemName == ee.Name && elemSize == getPipeDiameter(ee) && !alreadySetList.Contains(ee) && elemLength == tempLength/*&& ee.LookupParameter(spName).AsString() == ""*/)
                                {
                                    sameList.Add(ee);
                                    alreadySetList.Add(ee);
                                }
                            }

                            //分完組後針對所有相同的管去更改編號
                            if (sameList.Count() > 0)
                            {
                                using (Transaction trans = new Transaction(doc))
                                {
                                    trans.Start("寫入裁切編號");
                                    foreach (Element p in sameList)
                                    {
                                        Parameter pipeNum = p.LookupParameter(paraName1[3]);
                                        pipeNum.Set(keyToSet.ToString());
                                    }
                                    trans.Commit();
                                }
                            }
                            keyToSet += 1;

                            #region 放置多重品類標籤的功能，目前暫時不需要
                            //Element multiCateTag = findMultiCateTag(doc);
                            //using (Transaction trans = new Transaction(doc))
                            //{
                            //    trans.Start("放置多重品類標籤");
                            //    foreach (Element e in sameList)
                            //    {
                            //        Reference elemRefer = new Reference(e);
                            //        MEPCurve mepCrv = e as MEPCurve;
                            //        LocationCurve pipeLocate = mepCrv.Location as LocationCurve;
                            //        Curve pipeCrv = pipeLocate.Curve;
                            //        XYZ middlePt = pipeCrv.Evaluate(0.5, true);
                            //        IndependentTag fabricTag = IndependentTag.Create(doc, multiCateTag.Id, activeView.Id, elemRefer, false, TagOrientation.Horizontal, middlePt);
                            //        tagsToDelete.Add(fabricTag.Id);
                            //    }
                            //    trans.Commit();
                            //}
                            #endregion
                        }

                    }
                    catch
                    {
                        //MessageBox.Show("執行失敗");
                        //return Result.Failed;
                        //using (Transaction trans = new Transaction(doc))
                        //{
                        //    trans.Start("刪除暫時編號");
                        //    foreach (ElementId id in tagsToDelete)
                        //    {
                        //        doc.Delete(id);
                        //    }
                        //    trans.Commit();
                        //}
                        break;
                    }
                }
                transGroup.Assimilate();
            }


            return Result.Succeeded;
        }
        public Element findMultiCateTag(Document doc)
        {
            string tagetName = "M_裁切編號標籤";
            Element targetElement = null;
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MultiCategoryTags).WhereElementIsElementType();
            foreach (Element e in coll)
            {
                if (e.Name == tagetName) targetElement = e;
            }
            return targetElement;
        }
        public string getPipeDiameter(Element elem)
        {
            Document _doc = elem.Document;
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            double targetValue = 0.0;
            string targetValueStr = "";
            if (elem.Category.Id == pipe.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == duct.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == conduit.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == tray.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            return targetValueStr;
        }
    }
}
