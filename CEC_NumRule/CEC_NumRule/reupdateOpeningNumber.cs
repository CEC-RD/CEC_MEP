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
    [Transaction(TransactionMode.Manual)]
    class reupdateOpeningNumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            method m = new method();
            ViewPlan viewPlan = doc.ActiveView as ViewPlan;
            Level genLevel = viewPlan.GenLevel;
            List<string> castTypes = new List<string>()
            {
                "穿牆","穿版","穿樑"
            };
            List<BuiltInCategory> builts = new List<BuiltInCategory>() {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming };

            DialogResult dialogResult = MessageBox.Show("確定要將所有開口「全部重新編號」?", "CEC-MEP", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                using (Transaction trans2 = new Transaction(doc))
                {
                    trans2.Start("開口重新編號");
                    for (int i = 0; i < builts.Count; i++)
                    //foreach (BuiltInCategory built in builts)
                    {
                        int num = 1;

                        List<linkObject> objects = m.getLinkObjects(doc, builts[i],doc.ActiveView);
                        List<Element> targetList = m.getTargetElement(doc, castTypes[i],doc.ActiveView); //依照類別取得開口
                        if (targetList.Count == 0)
                        {
                            continue;
                        }
                        string numberPara = "開口編號";
                        //string systemPara = "系統別_映射";
                        foreach (linkObject ob in objects)
                        {
                            List<Element> castList = m.getCastInLinkObject(doc, ob, castTypes[i]);
                            if (castList.Count > 0)
                            {
                                foreach (Element cast in castList)
                                {
                                    Parameter para2 = cast.LookupParameter(numberPara);
                                    if (para2 == null) MessageBox.Show($"請檢查開口元件中是否有「{numberPara }」參數");

                                    para2.Set(num.ToString());
                                    num++;
                                }
                            }
                        }
                    }
                    trans2.Commit();
                }
                MessageBox.Show($"「{genLevel.Name}」中的開口重新編號完成!");
            }
            return Result.Succeeded;
        }
    }
}
