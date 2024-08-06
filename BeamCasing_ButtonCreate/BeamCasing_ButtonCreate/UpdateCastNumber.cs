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
    class UpdateCastNumber : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            IList<Element> levelList = findAllLevel(doc);
            int totalCount = 0;
            Level tempLevel;
            Counter.count += 1;
            try
            {
                #region 邏輯解釋
                //1.針對所有的穿樑套管編號
                //2.找到不在正常編號或是重複編號的套管，其餘的則放入既有套管編號(正常編號不需在編)
                //3.將原本不正常編號，或是多編的套管數字補上 (補位邏輯)
                #endregion
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("穿樑套管自動邊號");
                    ViewPlan viewPlan = doc.ActiveView as ViewPlan;
                    if (viewPlan == null)
                    {
                        MessageBox.Show("請在平面視圖中使用此功能");
                        return Result.Failed;
                    }

                    tempLevel = viewPlan.GenLevel as Level;
                    List<FamilyInstance> castList = findCastByLevel(tempLevel).OrderBy(x => getCastPt(x).X).ThenBy(x => getCastPt(x).Y).ToList();
                    //List<FamilyInstance> noNumList = new List<FamilyInstance>(); //沒有編號的套管
                    List<FamilyInstance> modifyList = new List<FamilyInstance>(); //編號需要修改的套管，沒有編號或編號錯誤的傢伙
                    List<string> toWrite = new List<string>(); //蒐集需要寫入的編號
                    List<string> numList = new List<string>();
                    List<string> existNumList = new List<string>();
                    int castCount = castList.Count();
                    string paraName = "開口編號";
                    //if (castCount == 0) continue;
                    totalCount += castCount;
                    //先蒐集一個正常的編號list
                    for (int i = 1; i < castCount + 1; i++)
                    {
                        string temp = "";
                        if (i < 10) temp = "0" + i.ToString();
                        else temp = i.ToString();
                        numList.Add(temp);
                    }
                    //針對既有編號的狀況檢查
                    foreach (FamilyInstance inst in castList)
                    {
                        string castNum = inst.LookupParameter(paraName).AsString();
                        //不在正常數字編號內，或是已經重複的套管，都必須修改
                        if (!numList.Contains(castNum) || existNumList.Contains(castNum))
                        {
                            modifyList.Add(inst);
                        }
                        else
                        {
                            //正確的數字才放入既有編號
                            existNumList.Add(castNum);
                        }
                    }
                    //不在既有編號中的正確數字，放入待寫名單中
                    foreach (string correct in numList)
                    {
                        if (!existNumList.Contains(correct))
                        {
                            toWrite.Add(correct);
                        }
                    }
                    //補上沒有寫入的數字
                    for (int i = 0; i < modifyList.Count(); i++)
                    {
                        modifyList[i].LookupParameter(paraName).Set(toWrite[i]);
                    }
                    if (totalCount == 0)
                    {
                        MessageBox.Show("模型中沒有任何穿牆套管，無法編號");
                        return Result.Failed;
                    }
                    trans.Commit();
                }
                MessageBox.Show($"「{tempLevel.Name}」中的穿樑套管編號完畢!");
            }
            catch
            {
                MessageBox.Show("執行失敗!");
                return Result.Failed;
            }

            return Result.Succeeded;
        }
        public IList<Element> findAllLevel(Document doc)
        {
            IList<Element> targetList = null;
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            ElementFilter level_Filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
            levelCollector.OfClass(typeof(Level)).WherePasses(level_Filter).WhereElementIsNotElementType();
            targetList = levelCollector.ToElements();
            return targetList;
        }
        public List<FamilyInstance> findCastByLevel(Level level)
        {
            List<FamilyInstance> targetList = new List<FamilyInstance>();
            ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);
            FilteredElementCollector castCollector = new FilteredElementCollector(level.Document).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(typeof(FamilyInstance)).WherePasses(levelFilter).WhereElementIsNotElementType();
            foreach (FamilyInstance inst in castCollector)
            {
                //針對checkName一定要確認是否為null，因為有些元件沒有此參數
                Parameter checkName = inst.Symbol.LookupParameter("API識別名稱");
                if (checkName != null && checkName.AsString().Contains("CEC-穿樑"))
                {
                    targetList.Add(inst);
                }
            }
            return targetList;
        }
        public XYZ getCastPt(FamilyInstance inst)
        {
            XYZ result = null;
            LocationPoint castLocate = inst.Location as LocationPoint;
            result = castLocate.Point;
            return result;
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
    }
}
