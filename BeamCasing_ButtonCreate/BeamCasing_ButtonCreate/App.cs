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
        //���ձN��Lbutton�[��{��TAB
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "��ٶ}�f";
        const string RIBBON_PANEL2 = "���CSD&SEM";
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
            //�Ыءu��ٮM�ޡv����
            List<RibbonPanel> panels = a.GetRibbonPanels(RIBBON_TAB); //�b���n�T�ORIBBON_TAB�b�o�椧�e�w�g�Q�Ы�
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

            //�ЫءuSEM&CSD�v����
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
            System.Drawing.Image image_CreateST = Properties.Resources.��ٮM��ICON�X��_ST;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            System.Drawing.Image image_CreateSTLink = Properties.Resources.��ٮM��ICON�X��_STlink;
            ImageSource imgSrc00 = GetImageSource(image_CreateSTLink);

            System.Drawing.Image image_Create = Properties.Resources.��ٮM��ICON�X��_RC;
            ImageSource imgSrc = GetImageSource(image_Create);

            System.Drawing.Image image_CreateLink = Properties.Resources.��ٮM��ICON�X��_RClink;
            ImageSource imgSrcLink = GetImageSource(image_CreateLink);

            System.Drawing.Image image_Update = Properties.Resources.��ٮM��ICON�X��_��s;
            ImageSource imgSrc2 = GetImageSource(image_Update);

            System.Drawing.Image image_UpdatePart = Properties.Resources.�ھڵ��Ͻd���s��ٸ�T_;
            ImageSource imgSrc22 = GetImageSource(image_UpdatePart);

            System.Drawing.Image image_SetUp = Properties.Resources.��ٮM��ICON�X��_�]�w;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);

            System.Drawing.Image image_Num = Properties.Resources.��ٮM��ICON�X��_�s��2;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_ReNum = Properties.Resources.��ٮM��ICON�X��_���s��2;
            ImageSource imgSrc5 = GetImageSource(image_ReNum);

            System.Drawing.Image image_Copy = Properties.Resources.�Ƭ�ٮM��ICON�X��_�ƻs;
            ImageSource imgSrc6 = GetImageSource(image_Copy);

            System.Drawing.Image image_Rect = Properties.Resources.��ٮM��ICON�X��_��ζ}�f;
            ImageSource imgSrc7 = GetImageSource(image_Rect);

            System.Drawing.Image image_RectLink = Properties.Resources.��ٮM��ICON�X��_��ζ}�flink;
            ImageSource imgSrc77 = GetImageSource(image_RectLink);

            System.Drawing.Image image_Multi = Properties.Resources.��ٮM��ICON�X��_�h�޶}�f;
            ImageSource imgSrc8 = GetImageSource(image_Multi);

            System.Drawing.Image image_MultiLink = Properties.Resources.��ٮM��ICON�X��_�h�޶}�flink;
            ImageSource imgSrc88 = GetImageSource(image_MultiLink);
            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_CastCreateST",
             "   ���c�}��   ",
             Assembly.GetExecutingAssembly().Location,
             "BeamCasing_ButtonCreate.CreateBeamCastSTV2"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
             );
            {
                btnData0.ToolTip = "�I��޻P�~�Ѽ٥ͦ���ٶ}�f";
                btnData0.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData00 = new PushButtonData(
 "MyButton_CastCreateSTLink",
 "   ���c�}��(�s��)   ",
 Assembly.GetExecutingAssembly().Location,
 "BeamCasing_ButtonCreate.CreateBeamCastLinkST"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
 );
            {
                btnData00.ToolTip = "�I��~�Ѻ޻P�~�Ѽ٥ͦ���ٶ}�f";
                btnData00.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnData00.LargeImage = imgSrc00;
            };

            PushButtonData btnData = new PushButtonData(
                "MyButton_CastCreate",
                "   RC�M��   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CreateBeamCastV2"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
                );
            {
                btnData.ToolTip = "�I��޻P�~�Ѽ٥ͦ���ٮM��";
                btnData.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnDataLink = new PushButtonData(
    "MyButton_CastCreateLink",
    "   RC�M��(�s��)   ",
    Assembly.GetExecutingAssembly().Location,
    "BeamCasing_ButtonCreate.CreateBeamCastLink"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
    );
            {
                btnDataLink.ToolTip = "�I��~�Ѻ޻P�~�Ѽ٥ͦ���ٮM��";
                btnDataLink.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnDataLink.LargeImage = imgSrcLink;
            };


            PushButtonData btnData2 = new PushButtonData(
                "MyButton_CastUpdate",
                "��s\n   ��ٸ�T   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CastInfromUpdateV4"
                );
            {
                btnData2.ToolTip = "�@���s��ٮM�޻P��ٶ}�f��T";
                btnData2.LongDescription = $"�̷ӥ��׳]�w������h�A��s��ٮM�޸�T (�������]�w��٭�h��i�ϥ�)({info})";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData22 = new PushButtonData(
                "MyButton_CastUpdatePart",
                "�~����s\n   ��ٸ�T   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.CastInfromUpdatePart"
                );
            {
                btnData22.ToolTip = "�̷ӥثe���Ͻd��A������s��ٮM�޸�T";
                btnData22.LongDescription = $"�̷ӥثe���Ͻd��A������s��ٮM�޸�T (�������]�w��٭�h��i�ϥ�)({info})";
                btnData22.LargeImage = imgSrc22;
            }

            PushButtonData btnData3 = new PushButtonData(
                "MyButton_CastSetUp",
                "�]�w\n   ��٭�h   ",
                Assembly.GetExecutingAssembly().Location,
                "BeamCasing_ButtonCreate.BeamCastSetUp"
                );
            {
                btnData3.ToolTip = "�]�w��٭�h����";
                btnData3.LongDescription = $"�̾ڱM�׻ݨD�A�]�w���ת���٭�h��T({info})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
    "MyButton_CastNum",
    "��ٮM��\n   �s��   ",
    Assembly.GetExecutingAssembly().Location,
    "BeamCasing_ButtonCreate.UpdateCastNumber"
    );
            {
                btnData4.ToolTip = "��ٮM�ަ۰ʽs��";
                btnData4.LongDescription = $"�ھڨC�h�Ӫ��}�f�ƶq�P��m�A�̧Ǧ۰ʱa�J�s���A�ĤG���W�J�s���ɫh�|���L�w�g��J�s�����M��({info})";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_ReNum",
"��ٮM��\n   ���s�s��   ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.ReUpdateCastNumber"
);
            {
                btnData5.ToolTip = "��ٮM�ޭ��s�s��";
                btnData5.LongDescription = $"�ھڨC�h�Ӫ��}�f�ƶq�A���s�a�J�s��({info})";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_CopyLinked",
"�ƻs�~��\n   ��ٮM��   ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CopyAllCast"
);
            {
                btnData6.ToolTip = "�ƻs�Ҧ��s���ҫ������M��";
                btnData6.LongDescription = $"�ƻs�Ҧ��s���ҫ������M�ޡA�H��SEM�}�f�s����({info})";
                btnData6.LargeImage = imgSrc6;
            }

            PushButtonData btnData7 = new PushButtonData(
"MyButton_CastRect",
"   ��μٶ}��    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CreateRectBeamCast"
);
            {
                btnData7.ToolTip = "�I��޻P�~�Ѽ٥ͦ���ٮM��";
                btnData7.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnData7.LargeImage = imgSrc7;
            }

            PushButtonData btnData77 = new PushButtonData(
"MyButton_CastRectLink",
"   ��μٶ}��(�s��)    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.CreateRectBeamCastLink"
);
            {
                btnData77.ToolTip = "�I��~�Ѻ޻P�~�Ѽ٥ͦ���ٮM��";
                btnData77.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~�Ѽ١A�ͦ���ٮM��({info})";
                btnData77.LargeImage = imgSrc77;
            }

            PushButtonData btnData8 = new PushButtonData(
"MyButton_MultiRect",
"   �h�޼ٶ}��    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.MultiBeamRectCast"
);
            {
                btnData8.ToolTip = "�I��h��޻P�~�Ѽ٥ͦ������}�f";
                btnData8.LongDescription = $"���I��ݭn�Ыت��ެq(�Ƽ�)�A�A�I����L���~����A�ͦ���٤�}�f({info})";
                btnData8.LargeImage = imgSrc8;
            }


            PushButtonData btnData88 = new PushButtonData(
"MyButton_MultiRectLink",
"   �h�޼ٶ}��(�s��)    ",
Assembly.GetExecutingAssembly().Location,
"BeamCasing_ButtonCreate.MultiBeamRectCastLink"
);
            {
                btnData88.ToolTip = "�I��h��~�Ѻ޻P�~�Ѽ٥ͦ������}�f";
                btnData88.LongDescription = $"���I��ݭn�Ыت��ެq(�Ƽ�)�A�A�I����L���~����A�ͦ���٤�}�f({info})";
                btnData88.LargeImage = imgSrc88;
            }

            //��s��ٸ�T(��s&�]�w)
            SplitButtonData setUpButtonData = new SplitButtonData("CastSetUpButton", "��ٮM�ާ�s");
            SplitButton splitButton1 = panel.AddItem(setUpButtonData) as SplitButton;
            PushButton button2 = splitButton1.AddPushButton(btnData2);
            button2 = splitButton1.AddPushButton(btnData22);
            button2 = splitButton1.AddPushButton(btnData3);

            //�Ыج�ٮM��(ST&RC)
            SplitButtonData STButtonData= new SplitButtonData("CreateCastST", "���c�}��");
            SplitButton splitButtonST = panel.AddItem(STButtonData) as SplitButton;
            PushButton STbutton = splitButtonST.AddPushButton(btnData0);
            STbutton = splitButtonST.AddPushButton(btnData00);

            SplitButtonData RCButtonData= new SplitButtonData("CreateCast","RC�M��");
            SplitButton splitButtonRC = panel.AddItem(RCButtonData) as SplitButton;
            PushButton RCbutton = splitButtonRC.AddPushButton(btnData);
            RCbutton = splitButtonRC.AddPushButton(btnDataLink);


            SplitButtonData rectCastButtonData = new SplitButtonData("RectCastButton", "��μٶ}�f");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button7 = splitButton.AddPushButton(btnData7);
            button7 = splitButton.AddPushButton(btnData77);
            button7 = splitButton.AddPushButton(btnData8);
            button7 = splitButton.AddPushButton(btnData88);

            //�ƻs�Ҧ��M��
            PushButton button6 = panel2.AddItem(btnData6) as PushButton;

            //��ٮM�޽s��(�s��&���s)
            SplitButtonData setNumButtonData = new SplitButtonData("CastSetNumButton", "��ٮM�޽s��");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button4 = splitButton2.AddPushButton(btnData4);
            button4 = splitButton2.AddPushButton(btnData5);
            //splitButton2.AddPushButton(btnData5);


            //�w�]Enabled���ӴN��true�A���ίS�O�]�w
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //���ưO�����x�s
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"��ٮM�޼Ҳզ@�Q�ϥΤF{Counter.count}��");
            return Result.Succeeded;
        }

        private BitmapSource GetImageSource(Image img)
        {
            //�s�@�@��function�M���ӳB�z�Ϥ�
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
