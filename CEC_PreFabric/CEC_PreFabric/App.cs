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
using System.Windows.Media; // for the graphics �ݤޥ�prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_PreFabric
{
    class App : IExternalApplication
    {
        //���ձN��Lbutton�[��{��TAB
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "�޽u�w��";
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
            //�Ыءu��ٮM�ޡv����
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

            System.Drawing.Image image_CreateISO = Properties.Resources.�w��ICON_����ISO��_svg_32;
            ImageSource imgSrc0 = GetImageSource(image_CreateISO);
            System.Drawing.Image image_CleanUpNumber = Properties.Resources.�w��ICON_�M���s��_svg_32;
            ImageSource imgSrc1 = GetImageSource(image_CleanUpNumber);
            System.Drawing.Image image_CleanTags = Properties.Resources.�w��ICON_�M������_svg_32;
            ImageSource imgSrc2 = GetImageSource(image_CleanTags);
            System.Drawing.Image image_Renumber = Properties.Resources.�w��ICON_���s�s��_svg_32;
            ImageSource imgSrc3 = GetImageSource(image_Renumber);

            //create all button data
            PushButtonData btnData0 = new PushButtonData(
                "MyButton_CreateFabISO",
                "�Ы�\n �w��ISO",
                Assembly.GetExecutingAssembly().Location,
                "CEC_PreFabric.CreateISO"
                );
            {
                btnData0.ToolTip = "�ؿ�ϰ첣�͹w��ISO��";
                btnData0.LongDescription = "�ؿ�Ʊ沣�͹w��ISO�PBOM���ϰ�A�ͦ�ISO�ϻPBOM��";
                btnData0.LargeImage = imgSrc0;
            }

            PushButtonData btnData1 = new PushButtonData(
                 "MyButton_ClenUpFabNum",
                 "�M���s��",
                 Assembly.GetExecutingAssembly().Location,
                 "CEC_PreFabric.CleanUpNumbers"
                 );
            {
                btnData1.ToolTip = "�M��ISO�Ϥ����ެq�丹";
                btnData1.LongDescription = "�Ф�����e���Ϧܱ��ק�s�����w�յ��ϡA�@��M���s��";
                btnData1.LargeImage = imgSrc1;
            }

            PushButtonData btnData2 = new PushButtonData(
                "MyButton_DeleteAllTags",
                "�M������",
                Assembly.GetExecutingAssembly().Location,
                "CEC_PreFabric.DeleteAllTags"
                );
            {
                btnData2.ToolTip = "�M��ISO�Ϥ����s������";
                btnData2.LongDescription = "�Ф�����e���Ϧܱ��M�����Ҫ��w�յ��ϡA�@��M������";
                btnData2.LargeImage = imgSrc2;
            }

            PushButtonData btnData3 = new PushButtonData(
    "MyButton_ISOReNumber",
    "���s�s��",
    Assembly.GetExecutingAssembly().Location,
    "CEC_PreFabric.ReNumber"
    );
            {
                btnData3.ToolTip = "�w��ISO�Ϥ����ެq�i�歫�s�s��";
                btnData3.LongDescription = "�Ф�����e���Ϧܱ����s�s�����w�յ��ϡA�妸���s�s��";
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
            //���ưO�����x�s
            ExcelLog log = new ExcelLog(Counter.count);
            log.userLog();
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
