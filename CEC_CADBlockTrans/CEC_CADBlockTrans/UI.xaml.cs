using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CEC_CADBlockTrans
{
    /// <summary>
    /// UI.xaml 的互動邏輯
    /// </summary>
    public partial class UI : Window
    {
        private readonly Document _doc;

        //private readonly UIApplication _uiApp;
        //private readonly Autodesk.Revit.ApplicationServices.Application _app;
        private readonly UIDocument _uiDoc;

        private readonly EventHandlerWithStringArg _mExternalMethodStringArg;
        private readonly EventHandlerWithWpfArg _mExternalMethodWpfArg;

        public UI(UIApplication uiApp, EventHandlerWithStringArg evExternalMethodStringArg,
            EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiDoc = uiApp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            Closed += MainWindow_Closed;
            InitializeComponent();

            //UI初始設定
            _mExternalMethodStringArg = evExternalMethodStringArg;
            _mExternalMethodWpfArg = eExternalMethodWpfArg;
            Method.getCADBlockList(this, _doc);
            Method.getSymbolCategory(this, _doc);
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            Close();
        }
        private void continueButtonClick(object sender, RoutedEventArgs e)
        {
            // Raise external event with this UI instance (WPF) as an argument
            //Method.DocumentInfo(this, _doc);
            //MessageBox.Show(this.BlockListBox.SelectedItems.Count.ToString());
            //foreach(var obj in BlockListBox.SelectedItems)
            //{
            //    Method.cadBlockCount(this, _doc, (ImportInstance)obj);
            //}

            //用來執行methodWrapper裡面的方法
            _mExternalMethodWpfArg.Raise(this);
        }
        private void cancelButtonClick(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            Close();
            return;
        }

        private void TextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            System.Windows.Controls.TextBox txt = sender as System.Windows.Controls.TextBox;

            //遮蔽非法按鍵
            if ((e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9) || e.Key == Key.Decimal)
            {
                if (txt.Text.Contains(".") && e.Key == Key.Decimal)
                {
                    e.Handled = true;
                    return;
                }
                e.Handled = false;
            }
            else if (((e.Key >= Key.D0 && e.Key <= Key.D9) || e.Key == Key.OemPeriod) && e.KeyboardDevice.Modifiers != ModifierKeys.Shift)
            {
                if (txt.Text.Contains(".") && e.Key == Key.OemPeriod)
                {
                    e.Handled = true;
                    return;
                }
                e.Handled = false;
            }
            else if (e.Key == Key.Delete || e.Key == Key.Back || e.Key == Key.Up || e.Key == Key.Right || e.Key == Key.Left || e.Key == Key.Down)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void categorySelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Method.getFamilyfromCategory(this, _doc);
        }

        private void familySelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Method.getSymbolsfromFamily(this, _doc);
        }
    }
}
