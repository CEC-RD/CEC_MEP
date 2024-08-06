using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Data;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CEC_Count
{
    public class EventHandlerWithStringArg : RevitEventWrapper<string>
    {
        /// <summary>
        /// The Execute override void must be present in all methods wrapped by the RevitEventWrapper.
        /// This defines what the method will do when raised externally.
        /// </summary>
        public override void Execute(UIApplication uiApp, string args)
        {
            // Do your processing here with "args"
            TaskDialog.Show("External Event", args);
        }
    }
    /// <summary>
    /// This is an example of of wrapping a method with an ExternalEventHandler using an instance of WPF
    /// as an argument. Any type of argument can be passed to the RevitEventWrapper, and therefore be used in
    /// the execution of a method which has to take place within a "Valid Revit API Context". This specific
    /// pattern can be useful for smaller applications, where it is convenient to access the WPF properties
    /// directly, but can become cumbersome in larger application architectures. At that point, it is suggested
    /// to use more "low-level" wrapping, as with the string-argument-wrapped method above.
    /// </summary>
    /// 
    public class EventHandlerWithWpfArg : RevitEventWrapper<CountingUI>
    {
        #region 更新和結束按鈕的部分
        public override void Execute(UIApplication uiApp, CountingUI ui)
        {
            //清空前一次的預設值
            Method.multiCheckDict = new Dictionary<ElementId, multiCheckItem>();
            Method.multiCheckItems = new List<multiCheckItem>();;

            List<Element> usefulList = new List<Element>() { };
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            List<CustomCate> cusCateLst = new List<CustomCate>();
            //Method m = new Method(uiApp);
            Method m = ui._method;

            //以勾選的品類作為計數單位
            int count = 0;
            bool viewFilterChecked = false;
            ui.Dispatcher.Invoke(() => ui.pBar.Value = 0);
            //ui.Dispatcher.Invoke(() => ui.pBar.Maximum = count);
            int selectedINdex = 0;
            ui.Dispatcher.Invoke(() => selectedINdex = ui.rvtLinkInstCombo.SelectedIndex);
            //排序以本機端檔案為第一個
            if (selectedINdex == 0)
            {
                Category massCate = Category.GetCategory(doc, BuiltInCategory.OST_Mass);
                CustomCate tempCate = new CustomCate()
                {
                    Name = massCate.Name,
                    Id = massCate.Id,
                    Cate = massCate,
                    Selected = true,
                    BuiltCate = BuiltInCategory.OST_Mass
                };
                cusCateLst.Add(tempCate);
                //ui.Dispatcher.Invoke(() => MessageBox.Show("選中第一個檔案"));
            }
            foreach (CustomCate cate in ui.mepCateList.Items)
            {
                if (cate.Selected == true)
                {
                    count++;
                    cusCateLst.Add(cate);
                }
            }
            string countMessage = $"共選擇了 {count} 個品類，正在進行數量計算...";
            ui.Dispatcher.Invoke(() => ui.selectHint.Text = countMessage);
            ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Black);
            LinkDoc linkDoc = null;
            //ui.Dispatcher.Invoke(() => targetDoc = ui.rvtLinkInstCombo.SelectedItem as LinkDoc);
            linkDoc = ui.rvtLinkInstCombo.SelectedItem as LinkDoc;
            ui.Dispatcher.Invoke(() => viewFilterChecked = ui.filterCheck.IsChecked.GetValueOrDefault());
            if (linkDoc != null)
            {
                Autodesk.Revit.DB.Transform targetTrans = linkDoc.Trans;
                //先取得量體總數，更新進度條上限
                Solid tempSolid = null;
                List<Element> massList = null;
                if (viewFilterChecked == true)
                {
                    ui.Dispatcher.Invoke(() => ui.selectHint.Text = "Loading....");
                    Solid solid = m.getSolidFromActiveView(doc, doc.ActiveView);
                    tempSolid = solid;
                    if (solid == null) ui.Dispatcher.Invoke(() => MessageBox.Show("找不到量體"));
                    ui.Dispatcher.Invoke(() => massList = m.getMassFromLinkDoc(linkDoc.Document, targetTrans, solid));
                }
                else
                {
                    ui.Dispatcher.Invoke(() => massList = m.getMassFromLinkDoc(linkDoc.Document, targetTrans));
                }
                ui.Dispatcher.Invoke(() => ui.pBar.Maximum = massList.Count());

                //針對量體總數進行干涉與進度更新
                using (TransactionGroup transGroup = new TransactionGroup(doc))
                {
                    transGroup.Start("分區數量計算");

                    //DirectShape ds =  m.createSolidFromBBox(doc.ActiveView) ;
                    foreach (Element e in massList)
                    {
                        List<Element> checkList = null;
                        ui.Dispatcher.Invoke(() => checkList = m.countByMass(cusCateLst, e, targetTrans));
                        ui.Dispatcher.Invoke(() => ui.pBar.Value += 1, System.Windows.Threading.DispatcherPriority.Background);
                    }
                    List<multiCheckItem> multiCheckItems = new List<multiCheckItem>();
                    ui.Dispatcher.Invoke(() => multiCheckItems = m.getMulitCheckItemFromDict(Method.multiCheckDict));
                    ui.Dispatcher.Invoke(() => ui.dataGrid1.ItemsSource = multiCheckItems);
                    m.removeUnuseElementPara(cusCateLst);//針對沒干涉到的進行重置

                    transGroup.Assimilate();
                }

                Task.Run(() =>
                    {
                        string completeMessage = $"計算完成!!";
                        ui.Dispatcher.Invoke(() => ui.selectHint.Text = completeMessage);
                        ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Blue);
                    });
                ui.Activate();
            }
            else if (linkDoc == null)
            {
                Task.Run(() =>
                {
                    string errorMessage = $"計算失敗，請確認是否選擇明確的量體來源模型!!";
                    ui.Dispatcher.Invoke(() => ui.selectHint.Text = errorMessage);
                    ui.Dispatcher.Invoke(() => ui.selectHint.Foreground = Brushes.Red);
                });
                ui.Activate();
            }
        }
        #endregion
    }
    public class ZoomHandlerWithWpfArg : RevitEventWrapper<CountingUI>
    {
        #region 負責查看按鈕的部分
        public override void Execute(UIApplication uiApp, CountingUI ui)
        {
            Method m = ui._method;
            m.zoomCurrentSelection(Method.multiCheckItems);
        }
        #endregion
    }
    public class UpdateHandlerWithWpfArg : RevitEventWrapper<CountingUI>
    {
        #region 負責更新按鈕的部分
        public override void Execute(UIApplication uiApp, CountingUI ui)
        {
            //string output = "";
            List<string> nameList = new List<string>();
            try
            {
                Method m = ui._method;
                m.multitemsReUpdate(Method.multiCheckItems);
            }
            catch { }
            //ui.Dispatcher.Invoke(() => ui.selectHint.Text = "按下更新按鈕!");
        }
        #endregion
    }
}
