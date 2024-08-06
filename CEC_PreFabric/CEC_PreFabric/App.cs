#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection; // for getting the assembly path
using System.Windows.Media; // for the graphics 需引用prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_PreFabric
{
    class App : IExternalApplication
    {
        //測試將其他button加到現有TAB
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "管線預組";
        public Result OnStartup(UIControlledApplication a)
        {
            RibbonPanel targetPanel = null;
            try
            {
                a.CreateRibbonTab(RIBBON_TAB);
            }
            catch (Exception)
            {
            }
            RibbonPanel panel = null;
            //創建「穿樑套管」頁籤
            List<RibbonPanel> panels = a.GetRibbonPanels(RIBBON_TAB);
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

            System.Drawing.Image image_CreateISO = Properties.Resources.預組ICON_產生ISO圖_svg_32;
            ImageSource imgSrc0 = GetImageSource(image_CreateISO);
            System.Drawing.Image image_CleanUpNumber = Properties.Resources.預組ICON_清除編號_svg_32;
            ImageSource imgSrc1 = GetImageSource(image_CleanUpNumber);
            System.Drawing.Image image_CleanTags = Properties.Resources.預組ICON_清除標籤_svg_32;
            ImageSource imgSrc2 = GetImageSource(image_CleanTags);
            System.Drawing.Image image_Renumber = Properties.Resources.預組ICON_重新編號_svg_32;
            ImageSource imgSrc3 = GetImageSource(image_Renumber);

            //create all button data
            PushButtonData btnData0 = new PushButtonData(
                "MyButton_CreateFabISO",
                "創建\n 預組ISO",
                Assembly.GetExecutingAssembly().Location,
                "CEC_PreFabric.CreateISO"
                );
            {
                btnData0.ToolTip = "框選區域產生預組ISO圖";
                btnData0.LongDescription = "框選希望產生預組ISO與BOM表的區域，生成ISO圖與BOM表";
                btnData0.LargeImage = imgSrc0;
            }

            PushButtonData btnData1 = new PushButtonData(
                 "MyButton_ClenUpFabNum",
                 "清除編號",
                 Assembly.GetExecutingAssembly().Location,
                 "CEC_PreFabric.CleanUpNumbers"
                 );
            {
                btnData1.ToolTip = "清除ISO圖中的管段邊號";
                btnData1.LongDescription = "請切換當前視圖至欲修改編號的預組視圖，一鍵清除編號";
                btnData1.LargeImage = imgSrc1;
            }

            PushButtonData btnData2 = new PushButtonData(
                "MyButton_DeleteAllTags",
                "清除標籤",
                Assembly.GetExecutingAssembly().Location,
                "CEC_PreFabric.DeleteAllTags"
                );
            {
                btnData2.ToolTip = "清除ISO圖中的編號標籤";
                btnData2.LongDescription = "請切換當前視圖至欲清除標籤的預組視圖，一鍵清除標籤";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData3 = new PushButtonData(
    "MyButton_ISOReNumber",
    "重新編號",
    Assembly.GetExecutingAssembly().Location,
    "CEC_PreFabric.ReNumber"
    );
            {
                btnData3.ToolTip = "針對ISO圖中的管段進行重新編號";
                btnData3.LongDescription = "請切換當前視圖至欲重新編號的預組視圖，批次重新編號";
                btnData3.LargeImage = imgSrc3;
            }

            //Add Pushbutton to Panel
            PushButton button = panel.AddItem(btnData0) as PushButton;
            PushButton button1 = panel.AddItem(btnData1) as PushButton;
            PushButton button2 = panel.AddItem(btnData2) as PushButton;
            PushButton button3 = panel.AddItem(btnData3) as PushButton;

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
    }
}
