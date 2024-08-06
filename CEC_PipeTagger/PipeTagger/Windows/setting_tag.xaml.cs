//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;

//namespace PipeTagger.Windows
//{
//    /// <summary>
//    /// setting_tag.xaml 的互動邏輯
//    /// </summary>
//    public partial class setting_tag : UserControl
//    {
//        public setting_tag()
//        {
//            InitializeComponent();
//        }
//    }
//}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PipeTagger.Windows
{
    /// <summary>
    /// NothingWindow.xaml 的互動邏輯
    /// </summary>
    public partial class setting_tag : Window
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;

        public setting_tag(UIDocument uIDoc)
        {
            uidoc = uIDoc;
            doc = uidoc.Document;
            InitializeComponent();
            Title = "管排標籤設定";

            //Image source
            PI_Image.Source = Common.UI_Tool.GetImageSource(Properties.Resources.PI_Crop_96);
            DT_Image.Source = Common.UI_Tool.GetImageSource(Properties.Resources.DT_Crop_96);
            CN_Image.Source = Common.UI_Tool.GetImageSource(Properties.Resources.CN_Crop_96);
            CT_Image.Source = Common.UI_Tool.GetImageSource(Properties.Resources.CT_Crop_96);
            Tag_Image.Source = Common.UI_Tool.GetImageSource(Properties.Resources.Tag_setting_dist);

            //1st_optio
            PI_1st_option.ItemsSource = Common.UI_Tool.Get_Tag_1st_options("管", doc);
            DT_C_1st_option.ItemsSource = Common.UI_Tool.Get_Tag_1st_options("圓風管", doc);
            DT_R_1st_option.ItemsSource = Common.UI_Tool.Get_Tag_1st_options("方風管", doc);
            CN_1st_option.ItemsSource =   Common.UI_Tool.Get_Tag_1st_options("電管", doc);
            CT_1st_option.ItemsSource =   Common.UI_Tool.Get_Tag_1st_options("電纜架", doc);

            PI_1st_option.SelectedItem = Properties.Settings.Default.PI_1st_Selected;
            DT_C_1st_option.SelectedItem = Properties.Settings.Default.DT_C_1st_Selected;
            DT_R_1st_option.SelectedItem = Properties.Settings.Default.DT_R_1st_Selected;
            CN_1st_option.SelectedItem = Properties.Settings.Default.CN_1st_Selected;
            CT_1st_option.SelectedItem = Properties.Settings.Default.CT_1st_Selected;
            
            //2nd_option
            //PI_2nd_option.ItemsSource =   Common.UI_Tool.Get_Tag_2nd_options(PI_1st_option.SelectedItem.ToString(), doc);
            //DT_C_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(DT_C_1st_option.SelectedItem.ToString(), doc);
            //DT_R_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(DT_R_1st_option.SelectedItem.ToString(), doc);
            //CN_2nd_option.ItemsSource =   Common.UI_Tool.Get_Tag_2nd_options(CN_1st_option.SelectedItem.ToString(), doc);
            //CT_2nd_option.ItemsSource =   Common.UI_Tool.Get_Tag_2nd_options(CT_1st_option.SelectedItem.ToString(), doc);
            
            PI_2nd_option.SelectedItem   = Properties.Settings.Default.PI_2nd_Selected;
            DT_C_2nd_option.SelectedItem = Properties.Settings.Default.DT_C_2nd_Selected;
            DT_R_2nd_option.SelectedItem = Properties.Settings.Default.DT_R_2nd_Selected;
            CN_2nd_option.SelectedItem   = Properties.Settings.Default.CN_2nd_Selected;
            CT_2nd_option.SelectedItem   = Properties.Settings.Default.CT_2nd_Selected;

            //Tag setting
            LenValue.Value = Properties.Settings.Default.label_length_mm;
            DistValue.Value = Properties.Settings.Default.label_dist_mm;
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {

            //if (PI_1st_option.SelectedItem == null ||
            //    DT_C_1st_option.SelectedItem == null ||
            //    DT_R_1st_option.SelectedItem == null ||
            //    CN_1st_option.SelectedItem == null ||
            //    CT_1st_option.SelectedItem == null) 
            //{
            //    MessageBox.Show("請確保每種管都有對應的標籤跟內容，或按「取消」後載入標籤。");
            //    return;
            //}


            if (PI_1st_option.SelectedItem != null)   { Properties.Settings.Default.PI_1st_Selected = PI_1st_option.SelectedItem.ToString(); }
            if (DT_C_1st_option.SelectedItem != null) { Properties.Settings.Default.DT_C_1st_Selected = DT_C_1st_option.SelectedItem.ToString();}
            if (DT_R_1st_option.SelectedItem != null) { Properties.Settings.Default.DT_R_1st_Selected = DT_R_1st_option.SelectedItem.ToString();}
            if (CN_1st_option.SelectedItem != null)   { Properties.Settings.Default.CN_1st_Selected = CN_1st_option.SelectedItem.ToString();}
            if (CT_1st_option.SelectedItem != null)   { Properties.Settings.Default.CT_1st_Selected = CT_1st_option.SelectedItem.ToString();}

            if (PI_2nd_option.SelectedItem != null)   { Properties.Settings.Default.PI_2nd_Selected = PI_2nd_option.SelectedItem.ToString();}
            if (DT_C_2nd_option.SelectedItem != null) { Properties.Settings.Default.DT_C_2nd_Selected = DT_C_2nd_option.SelectedItem.ToString();}
            if (DT_R_2nd_option.SelectedItem != null) { Properties.Settings.Default.DT_R_2nd_Selected = DT_R_2nd_option.SelectedItem.ToString();}
            if (CN_2nd_option.SelectedItem != null)   { Properties.Settings.Default.CN_2nd_Selected = CN_2nd_option.SelectedItem.ToString();}
            if (CT_2nd_option.SelectedItem != null)   { Properties.Settings.Default.CT_2nd_Selected = CT_2nd_option.SelectedItem.ToString();}
            
            Properties.Settings.Default.label_length_mm = (int)LenValue.Value;
            Properties.Settings.Default.label_dist_mm = (int)DistValue.Value;
            
            Properties.Settings.Default.Save();

            Close();
            return;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
            return;
        }

        #region SelectionChanged
        private void PI_1st_option_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PI_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(e.AddedItems[0].ToString(), doc);
            PI_2nd_option.SelectedIndex = 0;
        }

        private void DT_R_1st_option_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DT_R_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(e.AddedItems[0].ToString(), doc);
            DT_R_2nd_option.SelectedIndex = 0;
        }

        private void DT_C_1st_option_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DT_C_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(e.AddedItems[0].ToString(), doc);
            DT_C_2nd_option.SelectedIndex = 0;
        }

        private void CN_1st_option_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CN_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(e.AddedItems[0].ToString(), doc);
            CN_2nd_option.SelectedIndex = 0;
        }

        private void CT_1st_option_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CT_2nd_option.ItemsSource = Common.UI_Tool.Get_Tag_2nd_options(e.AddedItems[0].ToString(), doc);
            CT_2nd_option.SelectedIndex = 0;
        }

        #endregion


        //private void ShowSomeText(object sender, RoutedEventArgs e)
        //{
        //    Label_Name.Content = text.Text;
        //    Properties.Settings.Default.Test_text = text.Text; //Save for one time only
        //    Properties.Settings.Default.Save(); //Save for next open

        //    MessageBox.Show(text.Text); /*text is from the xaml file variable*/
        //    Close();
        //    return;
        //}
    }
}
