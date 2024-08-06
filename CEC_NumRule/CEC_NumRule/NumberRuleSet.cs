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
    public class NumberRuleSet : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,ref string message,ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            method m = new method();

            RuleSetting ruleUI = new RuleSetting(doc);
            ruleUI.pipeGrid.ItemsSource = m.GetCustomSystems(doc,BuiltInCategory.OST_PipingSystem);
            ruleUI.ductGrid.ItemsSource = m.GetCustomSystems(doc, BuiltInCategory.OST_DuctSystem);
            ruleUI.conduitGrid.ItemsSource = m.GetCustomSystems(doc, BuiltInCategory.OST_Conduit);
            ruleUI.trayGrid.ItemsSource = m.GetCustomSystems(doc, BuiltInCategory.OST_CableTray);
            //ruleUI.ductGrid.ItemsSource = ductSystemLst;
            //ruleUI.conduitGrid.ItemsSource = conduitTypeLst;
            //ruleUI.trayGrid.ItemsSource = cableTrayLst;
            ruleUI.ShowDialog();



            return Result.Succeeded;
        }
    }
}
