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
    //蒐集視圖內可見的柱、梁、牆、版套管，並進行系統更新
    [Transaction(TransactionMode.Manual)]
    class updateOpeningSystem : IExternalCommand
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

            //1.蒐集管、風管、電管、電纜架這四種BuiltinCategory
            //2.蒐集套管製作成Dictionary，key=套管，Value=各管
            //3.針對每一個套管蒐集它們的系統去蒐集Entity，如果沒有，則收集在「本機」所對應的Entity，並以List進行儲存
            //bool separateNumbering = false; //用來判斷是否需要分開蒐集套管-->目前尚未開發此功能
            List<string> castTypes = new List<string>()
            {
                "穿牆","穿版","穿樑"
            };

            List<BuiltInCategory> builts = new List<BuiltInCategory>() {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming };

            //List<Element> targetCast = m.getTargetElement(doc, castTypes);
            List<Element> targetCast = m.getTargetElement(doc, castTypes, doc.ActiveView);
            Dictionary<ElementId, List<Element>> castDict = m.getCastDict(doc, targetCast);
            List<string> paraNames = new List<string>()
            {
                "系統別_映射","不更新系統別"
            };
            foreach (ElementId id in castDict.Keys)
            {
                Element cast = doc.GetElement(id);
                FamilyInstance instance = cast as FamilyInstance;
                FamilySymbol symbol = instance.Symbol;
                Family family = symbol.Family;
                foreach(string st in paraNames)
                {
                    if (!m.checkPara(cast, st))
                    {
                        MessageBox.Show($"請檢查{family.Name}:{symbol.Name}中是否存在「{st}」參數");
                        return Result.Failed;
                    }
                }
            }

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("更新系統映射");
                m.reUpdateCastSystem(castDict, doc);
                trans.Commit();
            }
            MessageBox.Show($"「{genLevel.Name}」中新開口系統更新完成!");
            return Result.Succeeded;
        }
    }
}
