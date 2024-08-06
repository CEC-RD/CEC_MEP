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
using System.Windows.Media; // for the graphics �ݤޥ�prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_WallCast
{
    class App : IExternalApplication
    {
        //���ձN��Lbutton�[��{��TAB
        string info = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "����}�f";
        const string RIBBON_PANEL2 = "����CSD&SEM";
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
            System.Drawing.Image image_Update = Properties.Resources.����M��ICON�X��_��s_svg;
            ImageSource imgSrc0 = GetImageSource(image_Update);

            System.Drawing.Image image_UpdatePart = Properties.Resources.����M��ICON�X��_������s_svg;
            ImageSource imgSrc1 = GetImageSource(image_UpdatePart);

            System.Drawing.Image image_Create = Properties.Resources.����M��ICON�X��_��m_svg;
            ImageSource imgSrc = GetImageSource(image_Create);

            System.Drawing.Image image_CreateLink = Properties.Resources.����M��ICON�X��_��mlink_svg;
            ImageSource imgSrc00 = GetImageSource(image_CreateLink);

            System.Drawing.Image image_Copy = Properties.Resources.����M��ICON�X��_�ƻs�~��_svg;
            ImageSource imgSrc2 = GetImageSource(image_Copy);

            System.Drawing.Image image_SetUp = Properties.Resources.����M��ICON�X��_�s��_svg;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);


            System.Drawing.Image image_Num = Properties.Resources.����M��ICON�X��_���s�s��_svg;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_Rect = Properties.Resources.����M��ICON�X��_��}�f_svg;
            ImageSource imgSrc5 = GetImageSource(image_Rect);

            System.Drawing.Image image_RectLink = Properties.Resources.����M��ICON�X��_��}�flink_svg;
            ImageSource imgSrc55 = GetImageSource(image_RectLink);

            System.Drawing.Image image_MultiRect = Properties.Resources.����M��ICON�X��_�h�ޤ�}�f_svg;
            ImageSource imgSrc6 = GetImageSource(image_MultiRect);

            System.Drawing.Image image_MultiRectLink = Properties.Resources.����M��ICON�X��_�h�ޤ�}�flink_svg;
            ImageSource imgSrc66 = GetImageSource(image_MultiRectLink);


            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_WallCastUpdate",
             "��s\n   �����T   ",
             Assembly.GetExecutingAssembly().Location,
             "CEC_WallCast.WallCastUpdate"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
             );
            {
                btnData0.ToolTip = "�@���s����}�f��T";
                btnData0.LongDescription = $"�@���s����}�f��T({info})";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData00 = new PushButtonData(
 "MyButton_WallCastUpdatePart",
 "�~����s\n   �����T   ",
 Assembly.GetExecutingAssembly().Location,
 "CEC_WallCast.WallCastUpdatePart"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
 );
            {
                btnData00.ToolTip = "�̷ӥثe���Ͻd��A������s����}�f��T";
                btnData00.LongDescription = $"�̷ӥثe���Ͻd��A������s����}�f��T({info})";
                btnData00.LargeImage = imgSrc1;
            };


            PushButtonData btnData = new PushButtonData(
                "MyButton_WallCastCreate",
                "   ����M��   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CreateWallCastV2"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
                );
            {
                btnData.ToolTip = "�I��޻P�~����ͦ�����M��";
                btnData.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~����A�ͦ�����M��({info})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnDatalink = new PushButtonData(
    "MyButton_WallCastCreateLink",
    "   ����M��(�s��)   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.CreateWallCastLink"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
    );
            {
                btnDatalink.ToolTip = "�I��~�Ѻ޻P�~����ͦ�����M��";
                btnDatalink.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~����A�ͦ�����M��({info})";
                btnDatalink.LargeImage = imgSrc00;
            };

            PushButtonData btnData2 = new PushButtonData(
                "MyButton_WallCastCopy",
                "�ƻs�~��\n   ����M��   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CopyAllWallCast"
                );
            {
                btnData2.ToolTip = "�ƻs�Ҧ��s���ҫ������M��";
                btnData2.LongDescription = $"�ƻs�Ҧ��s���ҫ������M�ޡA�H��SEM�}�f�s����({info})";
                btnData2.LargeImage = imgSrc2;
            }


            PushButtonData btnData3 = new PushButtonData(
    "MyButton_WallCastNum",
    "����M��\n   �s��   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.UpdateWallCastNumber"
    );
            {
                btnData3.ToolTip = "����M�ަ۰ʽs��";
                btnData3.LongDescription = $"�ھڨC�h�Ӫ��}�f�ƶq�P��m�A�̧Ǧ۰ʱa�J�s���A�ĤG���W�J�s���ɫh�|���L�w�g��J�s�����M��({info})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
"MyButton_WallCastReNum",
"����M��\n   ���s�s��   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.ReUpdateWallCastNumber"
);
            {
                btnData4.ToolTip = "����M�ޭ��s�s��";
                btnData4.LongDescription = $"�ھڨC�h�Ӫ��}�f�ƶq�A���s�a�J�s��({info})";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_WallCastRect",
"   �諬��}�f   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCast"
);
            {
                btnData5.ToolTip = "�I��޻P�~����ͦ������}�f";
                btnData5.LongDescription = $"���I��ݭn�Ыت��ެq�A�A�I����L���~����A�ͦ������}�f({info})";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData55 = new PushButtonData(
"MyButton_WallCastRectLink",
"   �諬��}�f(�s��)   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCastLink"
);
            {
                btnData55.ToolTip = "�I��~�Ѻ޻P�~����ͦ������}�f";
                btnData55.LongDescription = $"���I��ݭn�Ыت��~�Ѻެq�A�A�I����L���~����A�ͦ������}�f({info})";
                btnData55.LargeImage = imgSrc55;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_WallCastRectMulti",
"   �h����}�f   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCast"
);
            {
                btnData6.ToolTip = "�I��~����P�h��ޥͦ������}�f";
                btnData6.LongDescription = $"���I��ݭn�Ыت��ެq(�Ƽ�)�A�A�I����L���~����A�ͦ������}�f({info})";
                btnData6.LargeImage = imgSrc6;
            }

            PushButtonData btnData66 = new PushButtonData(
"MyButton_WallCastRectMultiLink",
"   �h����}�f(�s��)   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCastLink"
);
            {
                btnData66.ToolTip = "�I��~����P�h��ޥͦ������}�f";
                btnData66.LongDescription = $"���I��ݭn�Ыت��ެq(�Ƽ�)�A�A�I����L���~����A�ͦ������}�f({info})";
                btnData66.LargeImage = imgSrc66;
            }


            //��s����M��
            SplitButtonData updateButtonData = new SplitButtonData("UpdateCast", "��s\n   �����T");
            SplitButton splitButton00 = panel.AddItem(updateButtonData) as SplitButton;
            PushButton button0 = splitButton00.AddPushButton(btnData0);
            button0 = splitButton00.AddPushButton(btnData00);
            //PushButton button0 = panel.AddItem(btnData0) as PushButton;

            //�Ыج���M��(��)
            SplitButtonData castButtonData = new SplitButtonData("WallCast", "����M��");
            SplitButton splitButton0 = panel.AddItem(castButtonData) as SplitButton;
            PushButton button = splitButton0.AddPushButton(btnData);
            button = splitButton0.AddPushButton(btnDatalink);

            //�Ыج���M��(��)
            SplitButtonData rectCastButtonData = new SplitButtonData("WallCastRect", "�諬��}�f");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button5 = splitButton.AddPushButton(btnData5);
            button5 = splitButton.AddPushButton(btnData55);
            button5 = splitButton.AddPushButton(btnData6);
            button5 = splitButton.AddPushButton(btnData66);

            //�ƻs�Ҧ��M��
            PushButton button2 = panel2.AddItem(btnData2) as PushButton;

            //��ٮM�޽s��(�s��&���s)
            SplitButtonData setNumButtonData = new SplitButtonData("WallCastSetNumButton", "����M�޽s��");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button3 = splitButton2.AddPushButton(btnData3);
            button3 = splitButton2.AddPushButton(btnData4);

            //�w�]Enabled���ӴN��true�A���ίS�O�]�w
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //���ưO�����x�s
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"����M�޼Ҳզ@�Q�ϥΤF{Counter.count}��");
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
