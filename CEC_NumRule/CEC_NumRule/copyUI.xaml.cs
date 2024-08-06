#region Namespaces
using System;
using System.Collections.Generic;
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
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;
using System.Threading;
#endregion

namespace CEC_NumRule
{
    /// <summary>
    /// copyUI.xaml 的互動邏輯
    /// </summary>
    public partial class copyUI : Window
    {
        public bool toCopy = false;
        public Document targetDoc = null;
        public copyUI()
        {
            InitializeComponent();
        }

        private void continueButton_Click(object sender, RoutedEventArgs e)
        {
            toCopy = true;
            Document tempDoc = this.linkDocCombo.SelectedItem as Document;
            targetDoc = tempDoc;
            Close();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            toCopy = false;
            Close();
            return;
        }
    }
}
