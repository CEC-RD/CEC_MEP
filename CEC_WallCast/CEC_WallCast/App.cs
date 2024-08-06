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
using System.Windows.Forms;
using System.Windows.Media; // for the graphics 需引用prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_WallCast
{
    class App : IExternalApplication
    {
        //測試將其他button加到現有TAB
        string info = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "穿牆開口";
        const string RIBBON_PANEL2 = "穿牆CSD&SEM";
        public Result OnStartup(UIControlledApplication a)
        {

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

            //創建「SEM&CSD」頁籤
            RibbonPanel panel2 = null;
            foreach (RibbonPanel pnl in panels)
            {
                if (pnl.Name == RIBBON_PANEL2)
                {
                    panel2 = pnl;
                    break;
                }
            }
            // couldn't find panel, create it
            if (panel2 == null)
            {
                panel2 = a.CreateRibbonPanel(RIBBON_TAB, RIBBON_PANEL2);
            }

            // get the image for the button
            System.Drawing.Image image_Update = Properties.Resources.穿牆套管ICON合集_更新_svg;
            ImageSource imgSrc0 = GetImageSource(image_Update);

            System.Drawing.Image image_UpdatePart = Properties.Resources.穿牆套管ICON合集_局部更新_svg;
            ImageSource imgSrc1 = GetImageSource(image_UpdatePart);

            System.Drawing.Image image_Create = Properties.Resources.穿牆套管ICON合集_放置_svg;
            ImageSource imgSrc = GetImageSource(image_Create);

            System.Drawing.Image image_CreateLink = Properties.Resources.穿牆套管ICON合集_放置link_svg;
            ImageSource imgSrc00 = GetImageSource(image_CreateLink);

            System.Drawing.Image image_Copy = Properties.Resources.穿牆套管ICON合集_複製外參_svg;
            ImageSource imgSrc2 = GetImageSource(image_Copy);

            System.Drawing.Image image_SetUp = Properties.Resources.穿牆套管ICON合集_編號_svg;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);


            System.Drawing.Image image_Num = Properties.Resources.穿牆套管ICON合集_重新編號_svg;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_Rect = Properties.Resources.穿牆套管ICON合集_方開口_svg;
            ImageSource imgSrc5 = GetImageSource(image_Rect);

            System.Drawing.Image image_RectLink = Properties.Resources.穿牆套管ICON合集_方開口link_svg;
            ImageSource imgSrc55 = GetImageSource(image_RectLink);

            System.Drawing.Image image_MultiRect = Properties.Resources.穿牆套管ICON合集_多管方開口_svg;
            ImageSource imgSrc6 = GetImageSource(image_MultiRect);

            System.Drawing.Image image_MultiRectLink = Properties.Resources.穿牆套管ICON合集_多管方開口link_svg;
            ImageSource imgSrc66 = GetImageSource(image_MultiRectLink);


            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_WallCastUpdate",
             "更新\n   穿牆資訊   ",
             Assembly.GetExecutingAssembly().Location,
             "CEC_WallCast.WallCastUpdate"//按鈕的全名-->要依照需要參照的command打入
             );
            {
                btnData0.ToolTip = "一鍵更新穿牆開口資訊";
                btnData0.LongDescription = $"一鍵更新穿牆開口資訊({info})";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData00 = new PushButtonData(
 "MyButton_WallCastUpdatePart",
 "居部更新\n   穿牆資訊   ",
 Assembly.GetExecutingAssembly().Location,
 "CEC_WallCast.WallCastUpdatePart"//按鈕的全名-->要依照需要參照的command打入
 );
            {
                btnData00.ToolTip = "依照目前視圖範圍，局部更新穿牆開口資訊";
                btnData00.LongDescription = $"依照目前視圖範圍，局部更新穿牆開口資訊({info})";
                btnData00.LargeImage = imgSrc1;
            };


            PushButtonData btnData = new PushButtonData(
                "MyButton_WallCastCreate",
                "   穿牆套管   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CreateWallCastV2"//按鈕的全名-->要依照需要參照的command打入
                );
            {
                btnData.ToolTip = "點選管與外參牆生成穿牆套管";
                btnData.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參牆，生成穿牆套管({info})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnDatalink = new PushButtonData(
    "MyButton_WallCastCreateLink",
    "   穿牆套管(連結)   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.CreateWallCastLink"//按鈕的全名-->要依照需要參照的command打入
    );
            {
                btnDatalink.ToolTip = "點選外參管與外參牆生成穿牆套管";
                btnDatalink.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參牆，生成穿牆套管({info})";
                btnDatalink.LargeImage = imgSrc00;
            };

            PushButtonData btnData2 = new PushButtonData(
                "MyButton_WallCastCopy",
                "複製外參\n   穿牆套管   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CopyAllWallCast"
                );
            {
                btnData2.ToolTip = "複製所有連結模型中的套管";
                btnData2.LongDescription = $"複製所有連結模型中的套管，以供SEM開口編號用({info})";
                btnData2.LargeImage = imgSrc2;
            }


            PushButtonData btnData3 = new PushButtonData(
    "MyButton_WallCastNum",
    "穿牆套管\n   編號   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.UpdateWallCastNumber"
    );
            {
                btnData3.ToolTip = "穿牆套管自動編號";
                btnData3.LongDescription = $"根據每層樓的開口數量與位置，依序自動帶入編號，第二次上入編號時則會略過已經填入編號的套管({info})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
"MyButton_WallCastReNum",
"穿牆套管\n   重新編號   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.ReUpdateWallCastNumber"
);
            {
                btnData4.ToolTip = "穿牆套管重新編號";
                btnData4.LongDescription = $"根據每層樓的開口數量，重新帶入編號({info})";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_WallCastRect",
"   方型牆開口   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCast"
);
            {
                btnData5.ToolTip = "點選管與外參牆生成穿牆方開口";
                btnData5.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參牆，生成穿牆方開口({info})";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData55 = new PushButtonData(
"MyButton_WallCastRectLink",
"   方型牆開口(連結)   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCastLink"
);
            {
                btnData55.ToolTip = "點選外參管與外參牆生成穿牆方開口";
                btnData55.LongDescription = $"先點選需要創建的外參管段，再點選其穿過的外參牆，生成穿牆方開口({info})";
                btnData55.LargeImage = imgSrc55;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_WallCastRectMulti",
"   多管牆開口   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCast"
);
            {
                btnData6.ToolTip = "點選外參牆與多支管生成穿牆方開口";
                btnData6.LongDescription = $"先點選需要創建的管段(複數)，再點選其穿過的外參牆，生成穿牆方開口({info})";
                btnData6.LargeImage = imgSrc6;
            }

            PushButtonData btnData66 = new PushButtonData(
"MyButton_WallCastRectMultiLink",
"   多管牆開口(連結)   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCastLink"
);
            {
                btnData66.ToolTip = "點選外參牆與多支管生成穿牆方開口";
                btnData66.LongDescription = $"先點選需要創建的管段(複數)，再點選其穿過的外參牆，生成穿牆方開口({info})";
                btnData66.LargeImage = imgSrc66;
            }


            //更新穿牆套管
            SplitButtonData updateButtonData = new SplitButtonData("UpdateCast", "更新\n   穿牆資訊");
            SplitButton splitButton00 = panel.AddItem(updateButtonData) as SplitButton;
            PushButton button0 = splitButton00.AddPushButton(btnData0);
            button0 = splitButton00.AddPushButton(btnData00);
            //PushButton button0 = panel.AddItem(btnData0) as PushButton;

            //創建穿牆套管(圓)
            SplitButtonData castButtonData = new SplitButtonData("WallCast", "穿牆套管");
            SplitButton splitButton0 = panel.AddItem(castButtonData) as SplitButton;
            PushButton button = splitButton0.AddPushButton(btnData);
            button = splitButton0.AddPushButton(btnDatalink);

            //創建穿牆套管(方)
            SplitButtonData rectCastButtonData = new SplitButtonData("WallCastRect", "方型牆開口");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button5 = splitButton.AddPushButton(btnData5);
            button5 = splitButton.AddPushButton(btnData55);
            button5 = splitButton.AddPushButton(btnData6);
            button5 = splitButton.AddPushButton(btnData66);

            //複製所有套管
            PushButton button2 = panel2.AddItem(btnData2) as PushButton;

            //穿樑套管編號(編號&重編)
            SplitButtonData setNumButtonData = new SplitButtonData("WallCastSetNumButton", "穿牆套管編號");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button3 = splitButton2.AddPushButton(btnData3);
            button3 = splitButton2.AddPushButton(btnData4);

            //預設Enabled本來就為true，不用特別設定
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //次數記錄後儲存
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"穿牆套管模組共被使用了{Counter.count}次");
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
