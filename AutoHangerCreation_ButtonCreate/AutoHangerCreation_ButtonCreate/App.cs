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
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace AutoHangerCreation_ButtonCreate
{
 class App : IExternalApplication
    {
        //用來創造按鈕 for 單管
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "管吊架";

        public Result OnStartup(UIControlledApplication a)
        {
            // get the ribbon tab
            try
            {
                a.CreateRibbonTab(RIBBON_TAB);
            }
            catch (Exception) { } //tab alreadt exists

            // get or create the panel
            RibbonPanel panel = null;
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

            // get the image for the button(放置單管吊架)
            System.Drawing.Image image_Single =Properties.Resources.單管_多管V2__轉換__02__96dpi;
            ImageSource imgSrc = GetImageSource(image_Single);

            //get the image for the button(放置多管吊架)
            System.Drawing.Image image_Multi =Properties.Resources.單管_多管V2__轉換__01__96dpi;
            ImageSource imgSrc2 = GetImageSource(image_Multi);

            //get the image for the button(調整吊架螺桿)
            System.Drawing.Image image_Adjust = Properties.Resources.吊桿長度調整_96DPI_01;
            ImageSource imgSrc3 = GetImageSource(image_Adjust);

            //get the image for the button(設定單管吊架)
            System.Drawing.Image image_SetUp = Properties.Resources.單管吊架設定_32pix;
            ImageSource imgSrc4 = GetImageSource(image_SetUp);

            //第三種做Button按鈕圖片的方法，參考官網對於button製作的描述
            //Uri uriImage = new Uri(@"D:\Dropbox (CHC Group)\工作人生\組內專案\02.Revit API開發\01.自動放置吊架\ICON\單管&多管V2 [轉換]-01).96dpi.png");
            //BitmapImage largeImage = new BitmapImage(uriImage);
            string assemblyInfo = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // create the button data
            PushButtonData btnData = new PushButtonData(
                "MyButton_Single",
                "創建\n   單管吊架   ",
                Assembly.GetExecutingAssembly().Location,
                "AutoHangerCreation_ButtonCreate.AddHangerByMouseLink"//按鈕的全名-->要依照需要參照的command打入
                );
            {
                btnData.ToolTip = "點選管段創建單管吊架";
                btnData.LongDescription = $"點選需要創建的管段，生成單管吊架({assemblyInfo})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnData2 = new PushButtonData("MyButton_Multi", "創建\n   多管吊架   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.MultiHangerCreationV3");
            {
                btnData2.ToolTip = "點選管段創建多管吊架";
                btnData2.LongDescription = $"點選需要創建的管段，生成多管吊架，單次最多選擇八支管({assemblyInfo})";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData3 = new PushButtonData("MyButton_ThreadAdjust", "調整\n   螺桿長度   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.HangerToFloorDist");
            {
                btnData3.ToolTip = "點選需要調整的吊架";
                btnData3.LongDescription = $"點選需要調整的吊架，調整螺桿長度連接至外參建築樓板({assemblyInfo})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData("MyButton_SetUp", "設定\n   單管吊架   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.PipeHangerSetUp");
            {
                btnData4.ToolTip = "設定吊架類型與間距";
                btnData4.LongDescription = $"設定自動放置單管吊架所需的吊架類型與間距，設定後才能使用單管吊架功能({assemblyInfo})";
                btnData4.LargeImage = imgSrc4;
            }

            //add the button to the ribbon
            PushButton button = panel.AddItem(btnData) as PushButton;
            PushButton button2 = panel.AddItem(btnData2) as PushButton;
            //PushButton button3 = panel.AddItem(btnData3) as PushButton;
            PushButton button4 = panel.AddItem(btnData4) as PushButton;

            //做完的button記得要Enable
            button.Enabled = true;
            button2.Enabled = true;
            //button3.Enabled = true;
            button4.Enabled = true;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //次數記錄後儲存
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"管吊架模組共被使用了{Counter.count}次");
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
