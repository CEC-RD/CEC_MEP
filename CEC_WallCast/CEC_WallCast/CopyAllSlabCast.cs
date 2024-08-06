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
    class CopyAllSlabCast : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public method m = new method();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            string messageOut = "";
            int copyNum = 0;
            try
            {
                string type = "穿版";
                //1.蒐集該檔案中所有的樓層元素
                //2.依照樓層篩選出實做的套管
                //3.以該樓層為基準，複製所有的套管
                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Dictionary<Transform, List<FamilyInstance>> linkCastDict = m.getLinkedCastDict(doc, type);
                if (linkCastDict.Keys.Count() == 0)
                {
                    MessageBox.Show("沒有實做的連結模型!");
                }
                List<FamilyInstance> targetInstList = m.findTargetElements(doc, type);
                List<bool> alreadyExist = new List<bool>();
                using (Transaction trans = new Transaction(doc))
                {
                    int deleteNum = 0;
                    trans.Start("複製所有外參檔中的穿版套管");
                    foreach (FamilyInstance cast in targetInstList)
                    {
                        if (!m.InstanceExistinLink(doc, cast, type))
                        {
                            doc.Delete(cast.Id);
                            deleteNum += 1;
                        }
                    }
                    MessageBox.Show($"本模型共刪除 {deleteNum} 個多餘的套管");
                    foreach (Transform t in linkCastDict.Keys)
                    {
                        //因為是以外參實體為單位去做實體蒐集
                        List<FamilyInstance> famList = linkCastDict[t];
                        List<ElementId> tempList = new List<ElementId>();
                        Document linkDoc = famList.First().Document;
                        foreach (FamilyInstance inst in famList)
                        {
                            //去除重複者的檢核機制
                            LocationPoint linkedInstLocate = inst.Location as LocationPoint;
                            XYZ linkedPt = linkedInstLocate.Point;
                            linkedPt = m.TransformPoint(linkedPt, t);
                            //準備要複製的這個跟已經有的去做比對，如果已經在模型中，則不複製，不在則需要複製
                            if (!m.InstanceExist(doc, linkedPt, type))
                            {
                                tempList.Add(inst.Id);
                            }
                        }
                        ICollection<ElementId> temp = tempList;
                        copyNum += temp.Count();
                        //有少才複製與顯示訊息
                        if (temp.Count > 0)
                        {
                            messageOut += $"來自 {linkDoc.Title} 的外參中少了 {tempList.Count()} 個套管\n";
                            CopyPasteOptions options = new CopyPasteOptions();
                            ElementTransformUtils.CopyElements(linkDoc, temp, doc, t, options);
                        }
                    }
                    trans.Commit();
                }

            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            messageOut += $"連結模型中的穿版套管已全數複製完成，共複製了 {copyNum} 個套管!\n";
            MessageBox.Show(messageOut);
            return Result.Succeeded;
        }
    }
}
