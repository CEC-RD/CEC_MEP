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

namespace CEC_PreFabric
{
    /// <summary>
    /// PreFabricUI.xaml 的互動邏輯
    /// </summary>
    public partial class PreFabricUI : Window
    {
        public PreFabricUI()
        {
            InitializeComponent();
        }
        //public bool cancelClick = false;
        private void ContinueButton_Click(object sender,RoutedEventArgs e)
        {
            Debug.WriteLine("Continue button was clicked");
            this.DialogResult = true;
            Close();
            return;
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            this.DialogResult = false;
            //cancelClick = true;
            Close();
            return;
        }

        private void TextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            TextBox txt = sender as TextBox;
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
    }
}
