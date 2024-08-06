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
    class CopyPartSlabCast : IExternalCommand
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
                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                Autodesk.Revit.DB.View activeView = doc.ActiveView;
                if (activeView.ViewType != ViewType.FloorPlan)
                {
                    MessageBox.Show("請於平面視圖中使用此功能");
                    return Result.Failed;
                }
                Level viewLevel = activeView.GenLevel;      
                Dictionary<Transform, List<FamilyInstance>> linkCastDict = m.getLinkedCastDictByLevel(doc, viewLevel,type);
                if (linkCastDict.Keys.Count() == 0)
                {
                    MessageBox.Show("沒有實作的連結模型");
                }
                List<FamilyInstance> targetInstList = m.findTargetElementsByLevels(doc, viewLevel,type);
                List<bool> alreadyExist = new List<bool>();
                using (Transaction trans = new Transaction(doc))
                {
                    int deleteNum = 0;
                    trans.Start($"複製{viewLevel.Name}中的穿版開口");
                    foreach (FamilyInstance cast in targetInstList)
                    {
                        if (!m.InstanceExistinLinkByLevel(doc, cast, viewLevel,type))
                        {
                            doc.Delete(cast.Id);
                            deleteNum += 1;
                        }
                    }
                    MessageBox.Show($"本模型共刪除 {deleteNum} 個多餘的開口");
                    foreach (Transform t in linkCastDict.Keys)
                    {
                        List<FamilyInstance> famList = linkCastDict[t];
                        List<ElementId> tempList = new List<ElementId>();
                        Document linkDoc = famList.First().Document;
                        foreach (FamilyInstance inst in famList)
                        {
                            //去除重複者的檢核機制
                            LocationPoint linkedInstLocate = inst.Location as LocationPoint;
                            XYZ linkedPt = linkedInstLocate.Point;
                            linkedPt = m.TransformPoint(linkedPt, t);
                            if (!m.InstanceExistByLevel(doc, linkedPt, viewLevel, type))
                            {
                                tempList.Add(inst.Id);
                            }
                        }
                        ICollection<ElementId> temp = tempList;
                        copyNum += temp.Count();
                        if (temp.Count() > 0)
                        {
                            messageOut += $"來自 {linkDoc.Title} 的外參於{viewLevel.Name}中少了 {tempList.Count()} 個開口\n";
                            CopyPasteOptions options = new CopyPasteOptions();
                            ElementTransformUtils.CopyElements(linkDoc, temp, doc, t, options);
                        }
                    }
                    messageOut += $"連結模型中位於{viewLevel.Name}的穿版開口已全數複製完成，共複製了 {copyNum} 個開口!\n";
                    trans.Commit();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            MessageBox.Show(messageOut);
            return Result.Succeeded;
        }
    }
}
