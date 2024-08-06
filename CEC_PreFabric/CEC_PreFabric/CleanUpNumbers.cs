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
#endregion

namespace CEC_PreFabric
{

    [Transaction(TransactionMode.Manual)]
    class CleanUpNumbers : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            string paraToCheck = "【預組】編號";
            if(activeView.ViewType != ViewType.ThreeD)
            {
                MessageBox.Show("本視圖並非3D視圖，請至3D視圖使用此功能");
                return Result.Failed;
            }
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
                BuiltInCategory[] builts = {
                                        BuiltInCategory.OST_PipeCurves,
                                        BuiltInCategory.OST_Conduit,
                                        BuiltInCategory.OST_DuctCurves,
                                        BuiltInCategory.OST_CableTray
                                       };
                List<ElementFilter> filters = new List<ElementFilter>();
                foreach (BuiltInCategory built in builts)
                {
                    ElementCategoryFilter filter = new ElementCategoryFilter(built);
                    filters.Add(filter);
                }
                LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
                collector.WherePasses(categoryFilter).WhereElementIsNotElementType();
                List<Element> usefulElement = new List<Element>();
                foreach (Element e in collector)
                {
                    if (e.LookupParameter(paraToCheck) != null)
                    {
                        usefulElement.Add(e);
                    }
                }
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("清除裁切編號");
                    if (usefulElement.Count() == 0)
                    {
                        MessageBox.Show($"視圖中沒有任何元件有名為「{paraToCheck}」的參數，無法清除編號");
                    }
                    else
                    {
                        foreach (Element e in usefulElement)
                        {
                            e.LookupParameter(paraToCheck).Set("");
                        }
                    }
                    trans.Commit();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
