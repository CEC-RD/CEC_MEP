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
#endregion


namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class UpdateSlabCastNumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            int totalCount = 0;
            Level tempLevel;
            Counter.count += 1;
            method m = new method();
            string keyName = "CEC-穿版";
            try
            {
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("穿版開口自動編號");
                    ViewPlan viewPlan = doc.ActiveView as ViewPlan;
                    if (viewPlan == null)
                    {
                        MessageBox.Show("請在平面視圖中使用此功能");
                        return Result.Failed;
                    }
                    //foreach (Element e in levelList)
                    //{
                    tempLevel = viewPlan.GenLevel as Level;
                    List<FamilyInstance> castList = m.findCastByLevel(tempLevel, keyName).OrderBy(x => m.getCastPt(x).X).ThenBy(x => m.getCastPt(x).Y).ToList();
                    List<FamilyInstance> modifyList = new List<FamilyInstance>(); //編號需要修改的套管，沒有編號或編號錯誤的傢伙
                    List<string> toWrite = new List<string>(); //蒐集需要寫入的編號
                    List<string> numList = new List<string>();
                    List<string> existNumList = new List<string>();
                    int castCount = castList.Count();
                    string paraName = "開口編號";
                    //if (castCount == 0) continue;
                    totalCount += castCount;
                    foreach (Element elem in castList)
                    {
                        if (!m.checkPara(elem, paraName))
                        {
                            MessageBox.Show($"套管中缺乏{paraName}參數欄位，請檢查!");
                            return Result.Failed;
                        }
                    }
                    //先蒐集一個正常的編號list
                    for (int i = 1; i < castCount + 1; i++)
                    {
                        string temp = "";
                        if (i < 10) temp = "0" + i.ToString();
                        else temp = i.ToString();
                        numList.Add(temp);
                    }
                    //針對既有編號的狀況檢查
                    foreach (FamilyInstance inst in castList)
                    {
                        string castNum = inst.LookupParameter(paraName).AsString();
                        //不在正常數字編號內，或是已經重複的套管，都必須修改
                        if (!numList.Contains(castNum) || existNumList.Contains(castNum))
                        {
                            modifyList.Add(inst);
                        }
                        else
                        {
                            //正確的數字才放入既有編號
                            existNumList.Add(castNum);
                        }
                    }
                    //不在既有編號中的正確數字，放入待寫名單中
                    foreach (string correct in numList)
                    {
                        if (!existNumList.Contains(correct))
                        {
                            toWrite.Add(correct);
                        }
                    }
                    //補上沒有寫入的數字
                    for (int i = 0; i < modifyList.Count(); i++)
                    {
                        modifyList[i].LookupParameter(paraName).Set(toWrite[i]);
                    }
                    //}
                    if (totalCount == 0)
                    {
                        MessageBox.Show("模型中沒有任何穿牆套管，無法編號");
                        return Result.Failed;
                    }
                    trans.Commit();
                }
                MessageBox.Show($"「{tempLevel.Name}」中的穿版開口編號完畢!");
            }
            catch
            {
                MessageBox.Show("執行失敗!");
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
