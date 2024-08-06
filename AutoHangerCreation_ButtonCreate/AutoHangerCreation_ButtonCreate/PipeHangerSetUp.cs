#region Namespaces
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;
using System.Linq;
using System;
using System.Drawing;
using System.Windows.Media.Imaging;
#endregion

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class PipeHangerSetUp : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;
                string output = "";


                //顯示設定視窗並選入更新值
                PipeHangerSetUpUI setUp_Window = new PipeHangerSetUpUI(commandData);
                setUp_Window.ShowDialog();
                if (PIpeHangerSetting.Default.DivideValueSelected == null || PIpeHangerSetting.Default.DivideValueSelected == PIpeHangerSetting.Default.FamilySelected)
                {
                    message = "單管吊架設定未完成 !!";
                    return Result.Failed;
                }
                if (PIpeHangerSetting.Default.DivideValueSelected == null ||  PIpeHangerSetting.Default.MultiHangerSelected==null)
                {
                    message = "多管吊架設定未完成 !!";
                    return Result.Failed;
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("錯誤，視窗無法顯示");
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
    class HangerTypeCollection
    {
        public List<string> getAllHangerNames(Document doc)
        {
            string paraName = "API識別名稱";
            string targetName = "吊架";
            //string[] hangerNames =
            //{
            //    "M_雙層管束_管附件",
            //    "M_葫蘆管束_管附件",
            //    "M_單管角鐵雙面孔吊架_管附件",
            //    "M_吊式PU保溫管墊吊架_管附件",
            //    "M_子母吊架_管附件",
            //    "M_UB束帶_管附件"
            //};
            List<string> hangerTypes = new List<string>();
            ElementFilter SymbolFilter = new ElementClassFilter(typeof(FamilySymbol));
            FilteredElementCollector hangerCollector = new FilteredElementCollector(doc);
            hangerCollector.WherePasses(SymbolFilter).ToElements();
            foreach (Element e in hangerCollector)
            {
                FamilySymbol tempSymbol = e as FamilySymbol;
                Parameter targetPara = tempSymbol.LookupParameter(paraName);
                if (targetPara != null && targetPara.AsString().Contains(targetName) && !hangerTypes.Contains(tempSymbol.Family.Name))
                {
                    hangerTypes.Add(tempSymbol.Family.Name);
                    continue;
                }
            }
            return hangerTypes;
        }
        public List<string> getMultiHangerNames(Document doc)
        {
            string paraName = "API識別名稱";
            string targetName = "多管";
            List<string> multiHangerTypes = new List<string>();
            ElementFilter SymbolFilter = new ElementClassFilter(typeof(FamilySymbol));
            FilteredElementCollector multiHangerCollector = new FilteredElementCollector(doc);
            multiHangerCollector.WherePasses(SymbolFilter).ToElements();
            foreach (Element e in multiHangerCollector)
            {
                FamilySymbol tempSymbol = e as FamilySymbol;
                Parameter targetPara = tempSymbol.LookupParameter(paraName);
                if (targetPara != null && targetPara.AsString().Contains(targetName) && !multiHangerTypes.Contains(tempSymbol.Family.Name))
                {
                    multiHangerTypes.Add(tempSymbol.Family.Name);
                    continue;
                }
            }
            return multiHangerTypes;

        }

        public BitmapImage getPreviewImage(Document doc, string familyName)
        {
            Family targetFamily = null;
            FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
            ElementFilter FamilyFilter = new ElementClassFilter(typeof(Family));
            familyCollector.WherePasses(FamilyFilter).ToElements();
            foreach (Family e in familyCollector)
            {
                if (e.Name == familyName)
                {
                    targetFamily = e;
                }
            }
            ICollection<ElementId> symbol_IDs = targetFamily.GetFamilySymbolIds();
            FamilySymbol tempSymbol = doc.GetElement(symbol_IDs.First()) as FamilySymbol; //取的第一個ID所代表的FamilySymbol
            ElementType type = tempSymbol as ElementType;
            System.Drawing.Size imgSize = new System.Drawing.Size(500, 500);
            Bitmap image = type.GetPreviewImage(imgSize);
            BitmapImage bitmapimage = new BitmapImage();

            using (MemoryStream memory = new MemoryStream())
            {
                image.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

            }
            return bitmapimage;
        }
    }

}
