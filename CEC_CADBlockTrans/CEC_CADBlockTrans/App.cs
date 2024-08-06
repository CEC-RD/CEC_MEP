#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media; // for the graphics 需引用prsentationCore
#endregion

namespace CEC_CADBlockTrans
{
    class App : IExternalApplication
    {
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "圖塊轉換";
        // class instance
        public static App ThisApp;

        public Result OnStartup(UIControlledApplication a)
        {
            //_mMyForm = null; // no dialog needed yet; the command will bring it
            //ThisApp = this; // static access to this application instance

            RibbonPanel targetPanel = null;
            // get the ribbon tab
            try
            {
                a.CreateRibbonTab(RIBBON_TAB);
            }
            catch (Exception) { } //tab alreadt exists
            RibbonPanel panel = null;
            //創建「穿樑套管」頁籤
            List<RibbonPanel> panels = a.GetRibbonPanels(RIBBON_TAB); //在此要確保RIBBON_TAB在這行之前已經被創建
            foreach (RibbonPanel pnl in panels)
            {
                if (pnl.Name == RIBBON_PANEL)
                {
                    panel = pnl;
                    break;
                }
            }
            // couldn't find panel, create it
            if (panel == null)
            {
                panel = a.CreateRibbonPanel(RIBBON_TAB, RIBBON_PANEL);
            }

            System.Drawing.Image image_CreateST = Properties.Resources.圖塊轉換icon_32pix;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            PushButtonData btnData0 = new PushButtonData(
             "MyButton_CADBlockTrans",
            "圖塊\n   批次轉換   ",
             Assembly.GetExecutingAssembly().Location,
            "CEC_CADBlockTrans.Command"//按鈕的全名-->要依照需要參照的command打入
            );
            {
                btnData0.ToolTip = "CAD圖塊批次轉換";
                btnData0.LongDescription = "CAD圖塊批次轉換";
                btnData0.LargeImage = imgSrc0;
            };

            PushButton button0 = panel.AddItem(btnData0) as PushButton;
            //button0.AvailabilityClassName = "CEC_CADBlockTrans.Availability";
            //設定按鈕僅能在平面視圖中看到


            // listeners/watchers for external events (if you choose to use them)
            a.ApplicationClosing += a_ApplicationClosing; //Set Application to Idling
            a.Idling += a_Idling;
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //次數記錄後儲存
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            return Result.Succeeded;
        }

        private BitmapSource GetImageSource(Image img)
        {
            //製作一個function專門來處理圖片
            BitmapImage bmp = new BitmapImage();
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                bmp.BeginInit();

                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = null;
                bmp.StreamSource = ms;

                bmp.EndInit();
            }
            return bmp;
        }

        void a_Idling(object sender, IdlingEventArgs e)
        {
        }
        /// <summary>
        /// What to do when the application is closing.)
        /// </summary>
        void a_ApplicationClosing(object sender, ApplicationClosingEventArgs e)
        {
        }
    }
}
