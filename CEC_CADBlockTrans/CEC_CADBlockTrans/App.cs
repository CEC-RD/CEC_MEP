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
using System.Windows.Media; // for the graphics �ݤޥ�prsentationCore
#endregion

namespace CEC_CADBlockTrans
{
    class App : IExternalApplication
    {
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "�϶��ഫ";
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

            System.Drawing.Image image_CreateST = Properties.Resources.�϶��ഫicon_32pix;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            PushButtonData btnData0 = new PushButtonData(
             "MyButton_CADBlockTrans",
            "�϶�\n   �妸�ഫ   ",
             Assembly.GetExecutingAssembly().Location,
            "CEC_CADBlockTrans.Command"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
            );
            {
                btnData0.ToolTip = "CAD�϶��妸�ഫ";
                btnData0.LongDescription = "CAD�϶��妸�ഫ";
                btnData0.LargeImage = imgSrc0;
            };

            PushButton button0 = panel.AddItem(btnData0) as PushButton;
            //button0.AvailabilityClassName = "CEC_CADBlockTrans.Availability";
            //�]�w���s�ȯ�b�������Ϥ��ݨ�


            // listeners/watchers for external events (if you choose to use them)
            a.ApplicationClosing += a_ApplicationClosing; //Set Application to Idling
            a.Idling += a_Idling;
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
