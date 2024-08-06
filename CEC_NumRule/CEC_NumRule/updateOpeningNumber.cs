#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
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

namespace CEC_NumRule
{
    //依照映射的系統別進行編號
    [Transaction(TransactionMode.Manual)]
    class updateOpeningNumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            method m = new method();
            ViewPlan viewPlan = doc.ActiveView as ViewPlan;
            Level genLevel = viewPlan.GenLevel;

            List<string> castTypes = new List<string>()
            {
                "穿牆","穿版","穿樑"
            };
            List<BuiltInCategory> builts = new List<BuiltInCategory>() {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming };

            //依不同類型的開口分類編號
            using (Transaction trans2 = new Transaction(doc))
            {
                trans2.Start("更新開口編號");
                for (int i = 0; i < builts.Count; i++)
                //foreach (BuiltInCategory built in builts)
                {
                    int num = 1;
                    string systemName = "";
                    List<linkObject> objects = m.getLinkObjects(doc, builts[i],doc.ActiveView);
                    MessageBox.Show("YA");
                    List<Element> targetList = m.getTargetElement(doc, castTypes[i],doc.ActiveView); //依照類別取得開口
                    int existMaxNum = 0;//針對不同品類的開口尋找最大編號值
                    string numberPara = "開口編號";
                    string systemPara = "系統別_映射";
                    foreach (Element e in targetList)
                    {
                        //if (systemName == "")
                        //{
                        //    systemName = e.LookupParameter(systemPara).AsString();
                        //}
                        //else if (systemName != e.LookupParameter(systemPara).AsString())
                        //{
                        //    systemName = e.LookupParameter(systemPara).AsString();
                        //    num = 1;
                        //}
                        Parameter para = e.LookupParameter(numberPara);
                        if (para == null) MessageBox.Show($"請檢查開口元件中是否有「{numberPara }」參數");
                        int tempNum = 0;
                        if (Int32.TryParse(para.AsString(), out tempNum))
                        {
                            tempNum = Int32.Parse(para.AsString());
                        }
                        //MessageBox.Show(tempNum.ToString());
                        if (tempNum > existMaxNum) existMaxNum = tempNum;
                    }
                    foreach (linkObject ob in objects)
                    {
                        List<Element> castList = m.getCastInLinkObject(doc, ob, castTypes[i]);
                        if (castList.Count > 0)
                        {
                            foreach (Element cast in castList)
                            {
                                int result = 0;
                                Parameter para2 = cast.LookupParameter(numberPara);
                                if (para2.AsString() == ""/* && existMaxNum != 1*/)//如果有套管還沒寫入數值，而且最大值不是1(表示先前已經寫入編號過)
                                {
                                    existMaxNum += 1;
                                    para2.Set(existMaxNum.ToString());
                                }
                                else if(Int32.TryParse(para2.AsString(),out result))
                                {
                                    para2.Set(num.ToString());
                                }
                                num++;
                            }
                        }
                    }
                }
                trans2.Commit();
            }
            MessageBox.Show($"「{genLevel.Name}」中的開口編號完成!");
            return Result.Succeeded;
        }
    }
}
