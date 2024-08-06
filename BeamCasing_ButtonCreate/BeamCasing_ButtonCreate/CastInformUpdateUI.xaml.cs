using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BeamCasing_ButtonCreate
{
    /// <summary>
    /// CastInformUpdateUI.xaml 的互動邏輯
    /// </summary>
    public partial class CastInformUpdateUI : Window
    {
        private ExternalCommandData CommandData;
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        public CastInformUpdateUI(ExternalCommandData commandData)
        {
            InitializeComponent();
            CommandData = commandData;
            uiapp = CommandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;

        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cotinue button was clicked.");
            Close();
            //UpdateCast updateCast = new UpdateCast(doc);
            //updateCast.findIntersectAndUpdate();
            //this.ProtectConflictListBox.ItemsSource = updateCast.Cast_tooClose;
            //this.TooCloseCastListBox.ItemsSource = updateCast.Cast_Conflict;
            //this.TooBigCastListBox.ItemsSource = updateCast.Cast_tooBig;
            //this.OtherCastListBox.ItemsSource = updateCast.Cast_OtherConfilct;
            //this.GriderCastListBox.ItemsSource = updateCast.Cast_BeamConfilct;
            return;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            Close();
            return;
        }

        private BitmapSource GetImageSource(System.Drawing.Image img)
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
