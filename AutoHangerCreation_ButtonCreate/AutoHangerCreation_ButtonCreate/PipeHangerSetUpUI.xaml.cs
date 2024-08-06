using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Drawing;

namespace AutoHangerCreation_ButtonCreate
{
    /// <summary>
    /// PipeHangerSetUpUI.xaml 的互動邏輯
    /// </summary>
    public partial class PipeHangerSetUpUI : Window
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        public BitmapImage catchImage { get; set; }
        public BitmapImage multiImage { get; set; }
        public string divideValue_setUp{ get; set; }
        public string targetFamily_setUp { get; set; }
        public PipeHangerSetUpUI(ExternalCommandData commandData)
        {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
            List<string> hangerFamilyList = new HangerTypeCollection().getAllHangerNames(doc);
            this.HangerTypeComboBox.ItemsSource = hangerFamilyList;
            List<string> multiHangerFamilyList = new HangerTypeCollection().getMultiHangerNames(doc);
            this.MultiHangerComboBox.ItemsSource = multiHangerFamilyList;

            //儲存預設值後在下次視窗開啟時顯示
            this.HangerTypeComboBox.SelectedItem = PIpeHangerSetting.Default.FamilySelected;
            this.MultiHangerComboBox.SelectedItem = PIpeHangerSetting.Default.MultiHangerSelected;
            this.DivideValue_TextBox.Text = PIpeHangerSetting.Default.DivideValueSelected;
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            //選擇後，寫入default setting檔
            PIpeHangerSetting.Default.DivideValueSelected = DivideValue_TextBox.Text;
            PIpeHangerSetting.Default.FamilySelected = HangerTypeComboBox.SelectedItem.ToString();
            PIpeHangerSetting.Default.MultiHangerSelected = MultiHangerComboBox.SelectedItem.ToString();
            PIpeHangerSetting.Default.Save();

            Debug.WriteLine("Ok button was clicked.");
            Close();
            return;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            Close();
            return;
        }

        private void HangerTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            catchImage = new HangerTypeCollection().getPreviewImage(doc, HangerTypeComboBox.SelectedItem.ToString());
            this.PreviewImage.Source = catchImage;
        }
        private void MultiTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            multiImage = new HangerTypeCollection().getPreviewImage(doc, MultiHangerComboBox.SelectedItem.ToString());
            this.MultiPreviewImage.Source = multiImage;
        }
    }
}
