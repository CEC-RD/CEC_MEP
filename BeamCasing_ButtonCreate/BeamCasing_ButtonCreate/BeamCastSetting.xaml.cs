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

namespace BeamCasing_ButtonCreate
{
    /// <summary>
    /// BeamCastSetting.xaml 的互動邏輯
    /// </summary>
    public partial class BeamCastSetting : Window
    {
        public BeamCastSetting()
        {
            InitializeComponent();
            //顯示大樑設定值
            DistC_ratio1.Text = BeamCast_Settings.Default.cD1_Ratio.ToString();
            ProtectC_ratio1.Text = BeamCast_Settings.Default.cP1_Ratio.ToString();
            ProtectC_min1.Text = BeamCast_Settings.Default.cP1_Min.ToString();
            SizeC_ratio1.Text = BeamCast_Settings.Default.cMax1_Ratio.ToString();
            SizeC_min1.Text = BeamCast_Settings.Default.cMax1_Max.ToString();
            DistR_ratio1.Text = BeamCast_Settings.Default.rD1_Ratio.ToString();
            ProtectR_ratio1.Text = BeamCast_Settings.Default.rP1_Ratio.ToString();
            ProtectR_min1.Text = BeamCast_Settings.Default.rP1_Min.ToString();
            SizeR_ratioD1.Text = BeamCast_Settings.Default.rMax1_RatioD.ToString();
            SizeR_ratioW1.Text = BeamCast_Settings.Default.rMax1_RatioW.ToString();

            //顯示小樑設定值
            DistC_ratio2.Text = BeamCast_Settings.Default.cD2_Ratio.ToString();
            ProtectC_ratio2.Text = BeamCast_Settings.Default.cP2_Ratio.ToString();
            ProtectC_min2.Text = BeamCast_Settings.Default.cP2_Min.ToString();
            SizeC_ratio2.Text = BeamCast_Settings.Default.cMax2_Ratio.ToString();
            SizeC_min2.Text = BeamCast_Settings.Default.cMax2_Max.ToString();
            DistR_ratio2.Text = BeamCast_Settings.Default.rD2_Ratio.ToString();
            ProtectR_ratio2.Text = BeamCast_Settings.Default.rP2_Ratio.ToString();
            ProtectR_min2.Text = BeamCast_Settings.Default.rP2_Min.ToString();
            SizeR_ratioD2.Text = BeamCast_Settings.Default.rMax2_RatioD.ToString();
            SizeR_ratioW2.Text = BeamCast_Settings.Default.rMax2_RatioW.ToString();
            LinkBox.IsChecked = BeamCast_Settings.Default.checkLink;
        }
        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cotinue button was clicked.");
            //按下確認後輸入設定值
            //大樑的預設值輸入
            BeamCast_Settings.Default.cD1_Ratio = Convert.ToDouble(DistC_ratio1.Text);
            BeamCast_Settings.Default.cP1_Ratio = Convert.ToDouble(ProtectC_ratio1.Text);
            BeamCast_Settings.Default.cP1_Min = Convert.ToDouble(ProtectC_min1.Text);
            BeamCast_Settings.Default.cMax1_Ratio = Convert.ToDouble(SizeC_ratio1.Text);
            BeamCast_Settings.Default.cMax1_Max = Convert.ToDouble(SizeC_min1.Text);
            BeamCast_Settings.Default.rD1_Ratio = Convert.ToDouble(DistR_ratio1.Text);
            BeamCast_Settings.Default.rP1_Ratio = Convert.ToDouble(ProtectR_ratio1.Text);
            BeamCast_Settings.Default.rP1_Min = Convert.ToDouble(ProtectR_min1.Text);
            BeamCast_Settings.Default.rMax1_RatioD = Convert.ToDouble(SizeR_ratioD1.Text);
            BeamCast_Settings.Default.rMax1_RatioW = Convert.ToDouble(SizeR_ratioW1.Text);

            //小樑的預設值輸入
            BeamCast_Settings.Default.cD2_Ratio = Convert.ToDouble(DistC_ratio2.Text);
            BeamCast_Settings.Default.cP2_Ratio = Convert.ToDouble(ProtectC_ratio2.Text);
            BeamCast_Settings.Default.cP2_Min = Convert.ToDouble(ProtectC_min2.Text);
            BeamCast_Settings.Default.cMax2_Ratio = Convert.ToDouble(SizeC_ratio2.Text);
            BeamCast_Settings.Default.cMax2_Max = Convert.ToDouble(SizeC_min2.Text);
            BeamCast_Settings.Default.rD2_Ratio = Convert.ToDouble(DistR_ratio2.Text);
            BeamCast_Settings.Default.rP2_Ratio = Convert.ToDouble(ProtectR_ratio2.Text);
            BeamCast_Settings.Default.rP2_Min = Convert.ToDouble(ProtectR_min2.Text);
            BeamCast_Settings.Default.rMax2_RatioD = Convert.ToDouble(SizeR_ratioD2.Text);
            BeamCast_Settings.Default.rMax2_RatioW = Convert.ToDouble(SizeR_ratioW2.Text);
            BeamCast_Settings.Default.Save();
            Close();
            return;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
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


        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //遮蔽中文輸入和非法字元貼上輸入
            TextBox textBox = sender as TextBox;
            TextChange[] change = new TextChange[e.Changes.Count];
            e.Changes.CopyTo(change, 0);

            int offset = change[0].Offset;
            if (change[0].AddedLength > 0)
            {
                double num = 0;
                if (!Double.TryParse(textBox.Text, out num))
                {
                    textBox.Text = textBox.Text.Remove(offset, change[0].AddedLength);
                    textBox.Select(offset, 0);
                }
            }
        }

        private void LinkBox_Checked(object sender, RoutedEventArgs e)
        {
            BeamCast_Settings.Default.checkLink  = true;
            BeamCast_Settings.Default.Save();
        }

        private void LinkBox_UnChecked(object sender, RoutedEventArgs e)
        {
            BeamCast_Settings.Default.checkLink = false;
            BeamCast_Settings.Default.Save();
        }
    }
}
