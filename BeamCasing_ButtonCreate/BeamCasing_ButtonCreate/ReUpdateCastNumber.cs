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
    class ReUpdateCastNumber : IExternalCommand
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
                DialogResult dialogResult = MessageBox.Show("確定要將穿樑套管「全部重新編號」?", "CEC-MEP", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("穿樑套管自動編號號");
                        ViewPlan viewPlan = doc.ActiveView as ViewPlan;
                        if (viewPlan == null)
                        {
                            MessageBox.Show("請在平面視圖中使用此功能");
                            return Result.Failed;
                        }
                        tempLevel = viewPlan.GenLevel as Level;
                        //foreach (Element e in levelList)
                        //{
                        int num = 1;
                        List<FamilyInstance> castList = findCastByLevel(tempLevel).OrderBy(x => getCastPt(x).X).ThenBy(x => getCastPt(x).Y).ToList();
                        List<FamilyInstance> noNumList = new List<FamilyInstance>();
                        List<FamilyInstance> modifyList = new List<FamilyInstance>();
                        //if (castList.Count == 0) continue;
                        totalCount += castList.Count;
                        foreach (FamilyInstance inst in castList)
                        {
                            //這邊要在做字串處理
                            string toWrite = "";
                            Parameter paraToSet = inst.LookupParameter("開口編號");
                            if (paraToSet != null)
                            {
                                if (num < 10)
                                {
                                    toWrite = "0" + num.ToString();
                                }
                                else
                                {
                                    toWrite = num.ToString();
                                }
                                paraToSet.Set(toWrite);
                                num++;
                            }
                        }
                        //}
                        if (totalCount == 0)
                        {
                            MessageBox.Show("模型中沒有任何穿樑套管，無法編號");
                            return Result.Failed;
                        }
                        trans.Commit();
                    }
                    MessageBox.Show($"「{tempLevel.Name}」中的穿樑套管重新編號完畢!");
                }
                else if (dialogResult == DialogResult.No)
                {
                }
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
            if (castCollector.Count() > 0)
            {
                foreach (FamilyInstance inst in castCollector)
                {
                    //針對checkName一定要確認是否為null，因為有些元件沒有此參數
                    Parameter checkName = inst.Symbol.LookupParameter("API識別名稱");
                    if (checkName != null && checkName.AsString().Contains("CEC-穿樑"))
                    {
                        targetList.Add(inst);
                    }
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

