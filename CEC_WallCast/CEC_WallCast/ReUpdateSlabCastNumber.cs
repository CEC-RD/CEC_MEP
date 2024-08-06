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


namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ReUpdateSlabCastNumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            int totalCount = 0;
            Level tempLevel;
            Counter.count += 1;
            method m = new method();
            string keyName = "CEC-穿版";
            try
            {
                DialogResult dialogResult = MessageBox.Show("確定要將穿牆套管「全部重新編號」?", "CEC-MEP", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("穿牆套管自動邊號");
                        ViewPlan viewPlan = doc.ActiveView as ViewPlan;
                        if (viewPlan == null)
                        {
                            MessageBox.Show("請在平面視圖中使用此功能");
                            return Result.Failed;
                        }
                        //forea
                        //foreach (Element e in levelList)
                        //{
                        tempLevel = viewPlan.GenLevel as Level;
                        int num = 1;
                        List<FamilyInstance> castList = m.findCastByLevel(tempLevel,keyName).OrderBy(x => m.getCastPt(x).X).ThenBy(x => m.getCastPt(x).Y).ToList();
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
                            MessageBox.Show("模型中沒有任何穿牆套管，無法編號");
                            return Result.Failed;
                        }
                        trans.Commit();
                    }
                    MessageBox.Show($"「{tempLevel.Name}」中的穿牆套管重新編號完畢!");
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
    }
}
