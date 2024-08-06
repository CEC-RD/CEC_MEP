#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using System.Threading;
#endregion

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class SlabCastUpdate : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //程式要做的事
            //1.蒐集所有的穿版套管
            //2.蒐集所有的外參版
            //3.更新穿版套管與版之間的關係-->是否穿牆、系統別與干涉管支數
            method m = new method();
            m.usefulLinkTrans = new List<Transform>();
            Counter.count += 1;
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw.Reset();//碼表歸零
            sw.Start();//碼表開始計時
            try
            {
                //準備東西
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;

                //找到所有穿版套管元件
                List<FamilyInstance> famList = m.findTargetElements(doc, "穿版");
                using (WallCastProgressUI progressView2 = new WallCastProgressUI(famList.Count))
                {
                    List<string> usefulParaName = new List<string> { "開口長", "干涉管數量", "系統別", "【原則】是否穿版", "不更新系統別" };
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("更新穿版套管資訊");
                        Dictionary<ElementId, List<Element>> slabCastDict = m.getCastDict(doc, "穿版");
                        int transCount = 0;
                        List<double> intersectVol = new List<double>();
                        foreach (FamilyInstance famInst in famList)
                        {
                            //檢查要更新的參數
                            foreach (string item in usefulParaName)
                            {
                                if (!m.checkPara(famInst, item))
                                {
                                    MessageBox.Show($"執行失敗，請檢查{famInst.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                                    return Result.Failed;
                                }
                            }
                            //檢查開口是否在版上
                            ElementId id = famInst.Id;
                            m.updateCastContent(doc, famInst);
                            int index = usefulParaName.IndexOf("【原則】是否穿版");
                            if (!slabCastDict.Keys.Contains(famInst.Id))
                            {
                                famInst.LookupParameter(usefulParaName[index]).Set("不符合");
                            }
                            else if (slabCastDict.Keys.Contains(famInst.Id))
                            {
                                famInst.LookupParameter(usefulParaName[index]).Set("OK");
                            }
                            if (progressView2.Update()) break;
                        }
                        trans.Commit();
                    }
                }
            }
            catch
            {
                MessageBox.Show("更新失敗");
            }
            sw.Stop();//碼錶停止
            double sec = Math.Round(sw.Elapsed.TotalMilliseconds / 1000, 2);
            string output = $"穿版開口資訊更新完成，共花費 {sec} 秒\n";
            MessageBox.Show(output + m.errorOutput);
            m.errorOutput = "";

            return Result.Succeeded;
        }
    }
}
