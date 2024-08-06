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
    class copyProjectSet : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            /*
             1. 先進行所有外參檔案的蒐集
             2. 選擇要以哪個檔案進行Entity的Copy
             */
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            List<Document> linkDocList = new List<Document>();
            FilteredElementCollector linkInstCollect = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType();
            foreach (Element e in linkInstCollect)
            {
                RevitLinkInstance tempInstance = e as RevitLinkInstance;
                if (tempInstance != null)
                {
                    Document tempDoc = tempInstance.GetLinkDocument();
                    linkDocList.Add(tempDoc);
                }
            }
            try
            {
                copyUI UI = new copyUI();
                UI.linkDocCombo.ItemsSource = linkDocList;
                UI.ShowDialog();

                if (UI.toCopy == true && UI.targetDoc != null)
                {
                    //Document targetDoc = linkDocList.First();
                    Document targetDoc = UI.targetDoc;

                    BuiltInCategory[] builts =
                    {
                BuiltInCategory.OST_PipingSystem,
                BuiltInCategory.OST_DuctSystem,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_CableTray
                };
                    List<ElementFilter> filters = new List<ElementFilter>();
                    foreach (BuiltInCategory built in builts)
                    {
                        ElementCategoryFilter filter = new ElementCategoryFilter(built);
                        filters.Add(filter);
                    }
                    LogicalOrFilter logicalOrFilter = new LogicalOrFilter(filters);

                    //建立資料圓
                    List<Element> sourceCollector = new FilteredElementCollector(targetDoc).WherePasses(logicalOrFilter).WhereElementIsElementType().ToList();
                    //建立要被寫入的檔案
                    List<Element> targetCollector = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsElementType().ToList();

                    //先找尋Schema，之後針對每個targetDoc要找尋的品類尋找Entity
                    Schema schema = Schema.Lookup(openingRuleScheme.civilSystemSchemaGUID);
                    if (schema != null)
                    {
                        using (TransactionGroup transGroup = new TransactionGroup(doc))
                        {
                            transGroup.Start("複製專案標準");
                            //針對每一個元素檢查Entity與移轉標準
                            foreach (Element e in sourceCollector)
                            {
                                Entity entity = e.GetEntity(schema);
                                string indexName = e.Name;
                                if (entity.IsValid())
                                {
                                    Element targetToSet = targetCollector.Find(x => x.Name == indexName);
                                    if (targetToSet != null)
                                    {
                                        openingRuleScheme.setEntityToElement(doc, schema, targetToSet, openingRuleScheme.retrieveDataFromElement(schema, e));
                                    }
                                }
                            }
                            transGroup.Assimilate();
                        }
                    }
                    else
                    {
                        MessageBox.Show("本專案尚未設定儲存值");
                    }
                    MessageBox.Show("移轉成功");
                }
            }
            catch
            {
                MessageBox.Show("移轉失敗");
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
