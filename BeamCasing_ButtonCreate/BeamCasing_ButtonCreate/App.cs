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
using System.Windows.Media; // for the graphics
using System.Windows.Media.Imaging;
using adWin = Autodesk.Windows;
#endregion

namespace BeamCasing_ButtonCreate
{

    class App : IExternalApplication
    {
        string info = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        //測試將其他button加到現有TAB
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "穿樑開口";
        const string RIBBON_PANEL2 = "穿樑CSD&SEM";
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
            System.Drawing.Image image_CreateST = Properties.Resources.穿樑套管ICON合集_ST;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            System.Drawing.Image image_CreateSTLink = Properties.Resources.穿樑套管ICON合集_STlink;
            ImageSource imgSrc00 = GetImageSource(image_CreateSTLink);

            System.Drawing.Image image_Create = Properties.Resources.穿樑套管ICON合集_RC;
            ImageSource imgSrc = GetImageSource(image_Create);

            System.Drawing.Image image_CreateLink = Properties.Resources.穿樑套管ICON合集_RClink;
            ImageSource imgSrcLink = GetImageSource(image_CreateLink);

            System.Drawing.Image image_Update = Properties.Resources.穿樑套管ICON合集_更新;
            ImageSource imgSrc2 = GetImageSource(image_Update);

            System.Drawing.Image image_UpdatePart = Properties.Resources.根據視圖範圍更新穿樑資訊_;
            ImageSource imgSrc22 = GetImageSource(image_UpdatePart);

            System.Drawing.Image image_SetUp = Properties.Resources.穿樑套管ICON合集_設定;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);

            System.Drawing.Image image_Num = Properties.Resources.穿樑套管ICON合集_編號2;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_ReNum = Properties.Resources.穿樑套管ICON合集_重編號2;
            ImageSource imgSrc5 = GetImageSource(image_ReNum);

            System.Drawing.Image image_Copy = Properties.Resources.副穿樑套管ICON合集_複製;
            ImageSource imgSrc6 = GetImageSource(image_Copy);

            System.Drawing.Image image_Rect = Properties.Resources.穿樑套管ICON合集_方形開口;
            ImageSource imgSrc7 = GetImageSource(image_Rect);

            System.Drawing.Image image_RectLink = Properties.Resources.穿樑套管ICON合集_方形開口link;
            ImageSource imgSrc77 = GetImageSource(image_RectLink);

            System.Drawing.Image image_Multi = Properties.Resources.穿樑套管ICON合集_多管開口;
            ImageSource imgSrc8 = GetImageSource(image_Multi);

            System.Drawing.Image image_MultiLink = Properties.Resources.穿樑套管ICON合集_多管開口link;
            ImageSource imgSrc88 = GetImageSource(image_MultiLink);
            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_CastCreateST",
             "   鋼構開孔   ",
             Assembly.GetExecutingAssembly().Location,
             "BeamCasing_ButtonCreate.CreateBeamCastSTV2"//按鈕的全名-->要依照需要參照的command打入
             );
            {
                btnData0.ToolTip = "點選管與外參樑生成穿樑開口";
                btnData0.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData00 = new PushButtonData(
 "MyButton_CastCreateSTLink",
 "   鋼構開孔(連結)   ",
 Assembly.GetExecutingAssembly().Location,
 "BeamCasing_ButtonCreate.CreateBeamCastLinkST"//按鈕的全名-->要依照需要參照的command打入
 );
            {
                btnData00.ToolTip = "點選外參管與外參樑生成穿樑開口";
                btnData00.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnData00.LargeImage = imgSrc00;
            };

            PushButtonData btnData = new PushButtonData(
                "MyButton_CastCreate",
                "   RC套管   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CreateBeamCastV2"//按鈕的全名-->要依照需要參照的command打入
                );
            {
                btnData.ToolTip = "點選管與外參樑生成穿樑套管";
                btnData.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnDataLink = new PushButtonData(
    "MyButton_CastCreateLink",
    "   RC套管(連結)   ",
    Assembly.GetExecutingAssembly().Location,
    "BeamCasing_ButtonCreate.CreateBeamCastLink"//按鈕的全名-->要依照需要參照的command打入
    );
            {
                btnDataLink.ToolTip = "點選外參管與外參樑生成穿樑套管";
                btnDataLink.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnDataLink.LargeImage = imgSrcLink;
            };


            PushButtonData btnData2 = new PushButtonData(
                "MyButton_CastUpdate",
                "更新\n   穿樑資訊   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CastInfromUpdateV4"
                );
            {
                btnData2.ToolTip = "一鍵更新穿樑套管與穿樑開口資訊";
                btnData2.LongDescription = $"依照本案設定的穿梁原則，更新穿樑套管資訊 (必須先設定穿樑原則方可使用)({info})";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData22 = new PushButtonData(
                "MyButton_CastUpdatePart",
                "居部更新\n   穿樑資訊   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CastInfromUpdatePart"
                );
            {
                btnData22.ToolTip = "依照目前視圖範圍，局部更新穿樑套管資訊";
                btnData22.LongDescription = $"依照目前視圖範圍，局部更新穿樑套管資訊 (必須先設定穿樑原則方可使用)({info})";
                btnData22.LargeImage = imgSrc22;
            }

            PushButtonData btnData3 = new PushButtonData(
                "MyButton_CastSetUp",
                "設定\n   穿樑原則   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.BeamCastSetUp"
                );
            {
                btnData3.ToolTip = "設定穿樑原則限制";
                btnData3.LongDescription = $"依據專案需求，設定本案的穿樑原則資訊({info})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
    "MyButton_CastNum",
    "穿樑套管\n   編號   ",
    Assembly.GetExecutingAssembly().Location,
    "BeamCasing_ButtonCreate.UpdateCastNumber"
    );
            {
                btnData4.ToolTip = "穿樑套管自動編號";
                btnData4.LongDescription = $"根據每層樓的開口數量與位置，依序自動帶入編號，第二次上入編號時則會略過已經填入編號的套管({info})";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_ReNum",
"穿樑套管\n   重新編號   ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.ReUpdateCastNumber"
);
            {
                btnData5.ToolTip = "穿樑套管重新編號";
                btnData5.LongDescription = $"根據每層樓的開口數量，重新帶入編號({info})";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_CopyLinked",
"複製外參\n   穿樑套管   ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CopyAllCast"
);
            {
                btnData6.ToolTip = "複製所有連結模型中的套管";
                btnData6.LongDescription = $"複製所有連結模型中的套管，以供SEM開口編號用({info})";
                btnData6.LargeImage = imgSrc6;
            }

            PushButtonData btnData7 = new PushButtonData(
"MyButton_CastRect",
"   方形樑開孔    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CreateRectBeamCast"
);
            {
                btnData7.ToolTip = "點選管與外參樑生成穿樑套管";
                btnData7.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnData7.LargeImage = imgSrc7;
            }

            PushButtonData btnData77 = new PushButtonData(
"MyButton_CastRectLink",
"   方形樑開孔(連結)    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CreateRectBeamCastLink"
);
            {
                btnData77.ToolTip = "點選外參管與外參樑生成穿樑套管";
                btnData77.LongDescription = $"先點選需要創建的管段，再點選其穿過的外參樑，生成穿樑套管({info})";
                btnData77.LargeImage = imgSrc77;
            }

            PushButtonData btnData8 = new PushButtonData(
"MyButton_MultiRect",
"   多管樑開孔    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.MultiBeamRectCast"
);
            {
                btnData8.ToolTip = "點選多支管與外參樑生成穿牆方開口";
                btnData8.LongDescription = $"先點選需要創建的管段(複數)，再點選其穿過的外參牆，生成穿樑方開口({info})";
                btnData8.LargeImage = imgSrc8;
            }


            PushButtonData btnData88 = new PushButtonData(
"MyButton_MultiRectLink",
"   多管樑開孔(連結)    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.MultiBeamRectCastLink"
);
            {
                btnData88.ToolTip = "點選多支外參管與外參樑生成穿牆方開口";
                btnData88.LongDescription = $"先點選需要創建的管段(複數)，再點選其穿過的外參牆，生成穿樑方開口({info})";
                btnData88.LargeImage = imgSrc88;
            }

            //更新穿樑資訊(更新&設定)
            SplitButtonData setUpButtonData = new SplitButtonData("CastSetUpButton", "穿樑套管更新");
            SplitButton splitButton1 = panel.AddItem(setUpButtonData) as SplitButton;
            PushButton button2 = splitButton1.AddPushButton(btnData2);
            button2 = splitButton1.AddPushButton(btnData22);
            button2 = splitButton1.AddPushButton(btnData3);

            //創建穿樑套管(ST&RC)
            SplitButtonData STButtonData= new SplitButtonData("CreateCastST", "鋼構開孔");
            SplitButton splitButtonST = panel.AddItem(STButtonData) as SplitButton;
            PushButton STbutton = splitButtonST.AddPushButton(btnData0);
            STbutton = splitButtonST.AddPushButton(btnData00);

            SplitButtonData RCButtonData= new SplitButtonData("CreateCast","RC套管");
            SplitButton splitButtonRC = panel.AddItem(RCButtonData) as SplitButton;
            PushButton RCbutton = splitButtonRC.AddPushButton(btnData);
            RCbutton = splitButtonRC.AddPushButton(btnDataLink);


            SplitButtonData rectCastButtonData = new SplitButtonData("RectCastButton", "方形樑開口");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button7 = splitButton.AddPushButton(btnData7);
            button7 = splitButton.AddPushButton(btnData77);
            button7 = splitButton.AddPushButton(btnData8);
            button7 = splitButton.AddPushButton(btnData88);

            //複製所有套管
            PushButton button6 = panel2.AddItem(btnData6) as PushButton;

            //穿樑套管編號(編號&重編)
            SplitButtonData setNumButtonData = new SplitButtonData("CastSetNumButton", "穿樑套管編號");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button4 = splitButton2.AddPushButton(btnData4);
            button4 = splitButton2.AddPushButton(btnData5);
            //splitButton2.AddPushButton(btnData5);


            //預設Enabled本來就為true，不用特別設定
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //次數記錄後儲存
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"穿樑套管模組共被使用了{Counter.count}次");
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
