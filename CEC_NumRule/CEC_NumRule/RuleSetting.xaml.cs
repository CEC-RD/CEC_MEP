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
    /// RuleSetting.xaml 的互動邏輯
    /// </summary>
    public partial class RuleSetting : Window
    {
        private Document _doc;

        public RuleSetting(Document doc)
        {
            InitializeComponent();
            _doc = doc;
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            //儲存新的系統縮寫
            string output = "";
            using (TransactionGroup transGroup = new TransactionGroup(_doc))
            {
                transGroup.Start("系統映射數據綁定");
                List<customSystem> unionList = new List<customSystem>();
                //先蒐集四個不同tab下的物件
                foreach (object obj in this.pipeGrid.Items)
                {
                    customSystem tempSystem = obj as customSystem;
                    unionList.Add(tempSystem);
                }
                foreach (object obj in this.ductGrid.Items)
                {
                    customSystem tempSystem = obj as customSystem;
                    unionList.Add(tempSystem);
                }
                foreach (object obj in this.conduitGrid.Items)
                {
                    customSystem tempSystem = obj as customSystem;
                    unionList.Add(tempSystem);
                }
                foreach (object obj in this.trayGrid.Items)
                {
                    customSystem tempSystem = obj as customSystem;
                    unionList.Add(tempSystem);
                }
                foreach (customSystem tempSystem in unionList)
                {
                    if (tempSystem.targetSystemName != null)
                    {
                        output += tempSystem.targetSystemName + "\n";
                        Schema schema = openingRuleScheme.getCivilSystemSchema();
                        openingRuleScheme.setEntityToElement(_doc, schema, tempSystem.systemType, tempSystem.targetSystemName);
                    }
                }
                transGroup.Assimilate();
            }
            Close();
            //MessageBox.Show(output);
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("Cancel button was clicked.");
            Close();
            return;
        }
    }
}
