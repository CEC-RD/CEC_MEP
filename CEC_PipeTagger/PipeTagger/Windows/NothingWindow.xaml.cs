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
    public partial class NothingWindow : Window
    {
        public UIDocument uidoc { get; }
        public Document doc { get; }

        public NothingWindow(UIDocument uIDoc)
        {
            uidoc = uIDoc;
            doc = uidoc.Document;
            InitializeComponent();
            Title = "Test view";
        }
        private void ShowSomeText(object sender, RoutedEventArgs e)
        {
            Label_Name.Content = text.Text;
            Properties.Settings.Default.Test_text = text.Text; //Save for one time only
            Properties.Settings.Default.Save(); //Save for next open

            MessageBox.Show(text.Text); /*text is from the xaml file variable*/
            Close();
            return;
        }
        }
    }
