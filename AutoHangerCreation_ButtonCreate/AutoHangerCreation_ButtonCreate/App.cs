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
        //�Ψӳгy���s for ���
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "�ަQ�[";

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

            // get the image for the button(��m��ަQ�[)
            System.Drawing.Image image_Single =Properties.Resources.���_�h��V2__�ഫ__02__96dpi;
            ImageSource imgSrc = GetImageSource(image_Single);

            //get the image for the button(��m�h�ަQ�[)
            System.Drawing.Image image_Multi =Properties.Resources.���_�h��V2__�ഫ__01__96dpi;
            ImageSource imgSrc2 = GetImageSource(image_Multi);

            //get the image for the button(�վ�Q�[����)
            System.Drawing.Image image_Adjust = Properties.Resources.�Q����׽վ�_96DPI_01;
            ImageSource imgSrc3 = GetImageSource(image_Adjust);

            //get the image for the button(�]�w��ަQ�[)
            System.Drawing.Image image_SetUp = Properties.Resources.��ަQ�[�]�w_32pix;
            ImageSource imgSrc4 = GetImageSource(image_SetUp);

            //�ĤT�ذ�Button���s�Ϥ�����k�A�Ѧҩx�����button�s�@���y�z
            //Uri uriImage = new Uri(@"D:\Dropbox (CHC Group)\�u�@�H��\�դ��M��\02.Revit API�}�o\01.�۰ʩ�m�Q�[\ICON\���&�h��V2 [�ഫ]-01).96dpi.png");
            //BitmapImage largeImage = new BitmapImage(uriImage);
            string assemblyInfo = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // create the button data
            PushButtonData btnData = new PushButtonData(
                "MyButton_Single",
                "�Ы�\n   ��ަQ�[   ",
                Assembly.GetExecutingAssembly().Location,
                "AutoHangerCreation_ButtonCreate.AddHangerByMouseLink"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
                );
            {
                btnData.ToolTip = "�I��ެq�Ыس�ަQ�[";
                btnData.LongDescription = $"�I��ݭn�Ыت��ެq�A�ͦ���ަQ�[({assemblyInfo})";
                btnData.LargeImage = imgSrc;
            };

            PushButtonData btnData2 = new PushButtonData("MyButton_Multi", "�Ы�\n   �h�ަQ�[   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.MultiHangerCreationV3");
            {
                btnData2.ToolTip = "�I��ެq�Ыئh�ަQ�[";
                btnData2.LongDescription = $"�I��ݭn�Ыت��ެq�A�ͦ��h�ަQ�[�A�榸�̦h��ܤK���({assemblyInfo})";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData3 = new PushButtonData("MyButton_ThreadAdjust", "�վ�\n   �������   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.HangerToFloorDist");
            {
                btnData3.ToolTip = "�I��ݭn�վ㪺�Q�[";
                btnData3.LongDescription = $"�I��ݭn�վ㪺�Q�[�A�վ�������׳s���ܥ~�ѫؿv�ӪO({assemblyInfo})";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData("MyButton_SetUp", "�]�w\n   ��ަQ�[   ", Assembly.GetExecutingAssembly().Location, "AutoHangerCreation_ButtonCreate.PipeHangerSetUp");
            {
                btnData4.ToolTip = "�]�w�Q�[�����P���Z";
                btnData4.LongDescription = $"�]�w�۰ʩ�m��ަQ�[�һݪ��Q�[�����P���Z�A�]�w��~��ϥγ�ަQ�[�\��({assemblyInfo})";
                btnData4.LargeImage = imgSrc4;
            }

            //add the button to the ribbon
            PushButton button = panel.AddItem(btnData) as PushButton;
            PushButton button2 = panel.AddItem(btnData2) as PushButton;
            //PushButton button3 = panel.AddItem(btnData3) as PushButton;
            PushButton button4 = panel.AddItem(btnData4) as PushButton;

            //������button�O�o�nEnable
            button.Enabled = true;
            button2.Enabled = true;
            //button3.Enabled = true;
            button4.Enabled = true;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            //���ưO�����x�s
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
            //MessageBox.Show($"�ަQ�[�Ҳզ@�Q�ϥΤF{Counter.count}��");
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
