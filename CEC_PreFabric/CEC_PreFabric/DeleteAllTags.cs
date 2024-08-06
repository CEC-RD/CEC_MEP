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
    class DeleteAllTags : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            string tagName = "裁切編號";
            if (activeView.ViewType != ViewType.ThreeD)
            {
                MessageBox.Show("本視圖並非3D視圖，請至3D視圖使用此功能");
                return Result.Failed;
            }
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id);
                ElementCategoryFilter categoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_MultiCategoryTags);
                collector.WherePasses(categoryFilter).WhereElementIsNotElementType();
                List<Element> targetList = new List<Element>();
                foreach (Element e in collector)
                {
                    if (e.Name.Contains(tagName))
                    {
                        targetList.Add(e);
                    }
                }
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("刪除裁切編號標籤");
                    if (targetList.Count() == 0)
                    {
                        MessageBox.Show("此視圖中並沒有關於裁切編號的標籤，無法刪除");
                    }
                    else
                    {
                        foreach (Element ee in targetList)
                        {
                            doc.Delete(ee.Id);
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
