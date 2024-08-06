using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Autodesk.Revit.DB.Structure;
using System;
using Autodesk.Revit.UI.Selection;
using System.Linq;

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class HangerToFloorDist : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Autodesk.Revit.UI.Selection.ISelectionFilter pipeAccess_Filter = new PipeAccessoryFilter();

                //點選要更改牙桿長度的吊架
                List<Element> pickAccessory = new List<Element>();
                Document doc = uidoc.Document;
                IList<Reference> pickAccessory_Refer = uidoc.Selection.PickObjects(ObjectType.Element, pipeAccess_Filter, $"請選整欲調整牙桿長度的吊架");

                foreach (Reference reference in pickAccessory_Refer)
                {
                    Element element = doc.GetElement(reference.ElementId);
                    pickAccessory.Add(element);
                }

                //ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
                //ICollection<Element> elementList = new List<Element>();

                Transaction trans = new Transaction(doc);
                trans.Start("調整螺牙長度");
                //將選到的物件寫進iList中
                try
                {
                    foreach (Element elem in pickAccessory)
                    {
                        FamilyInstance instance = elem as FamilyInstance;
                        double threadLength = CalculateDist_upperLevel(doc, instance);
                        instance.LookupParameter("PipeCenterToFloor").Set(threadLength);
                    }

                    trans.Commit();
                    MessageBox.Show("螺桿長度調整完畢!");
                }
                catch
                {
                    //MessageBox.Show("請檢查是否有在剖面框中使用此功能，或是否有選到上方沒有樓板或存在樑中的吊架!!");
                    MessageBox.Show("請檢查預設3D視圖中的剖面框是啟用，且範圍中是否包括上方支撐點!!");
                    trans.RollBack();
                    return Result.Failed;
                }
            }
            catch
            {
                return Result.Failed;
            }
            Counter.count += 1;
            return Result.Succeeded;
        }

        private double CalculateDist_upperLevel(Document doc, FamilyInstance hanger)
        {
            //利用ReferenceIntersector回傳吊架location point 和上層樓板之間的距離
            //Find a 3D view to use for ReferenceIntersector constructor
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            Func<View3D, bool> isNotTemplate = v3 => !(v3.IsTemplate);
            //View3D view3D = collector.OfClass(typeof(View3D)).Cast<View3D>().First<View3D>(isNotTemplate);
            View3D view3D = collector.OfClass(typeof(View3D)).Cast<View3D>().First(isNotTemplate);
            //Find the locationiPoint of Hanger as the start point
            LocationPoint hangerLocation = hanger.Location as LocationPoint;
            XYZ startLocation = hangerLocation.Point;

            //Project in the positive Z direction on to the floor
            XYZ rayDirectioin = new XYZ(0, 0, 1);
            //ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
            ElementFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            ElementFilter filter1 = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            LogicalOrFilter orFilter = new LogicalOrFilter(filter, filter1);
            //ElementClassFilter filter1 = new ElementClassFilter(t)
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(orFilter, FindReferenceTarget.Face, view3D);

            //FindReferencesInRevitLinks=true 打開對於外參的測量
            referenceIntersector.FindReferencesInRevitLinks = true;
            ReferenceWithContext referenceWithContext = referenceIntersector.FindNearest(startLocation, rayDirectioin);

            Reference reference = referenceWithContext.GetReference();
            XYZ intersection = reference.GlobalPoint;

            double dist = startLocation.DistanceTo(intersection);
            return dist;
        }

        public class PipeAccessoryFilter : Autodesk.Revit.UI.Selection.ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                FamilyInstance inst = element as FamilyInstance;
                FamilySymbol symbol = inst.Symbol;
                Parameter targetPara = symbol.LookupParameter("API識別名稱");
                //string elementName = symbol.FamilyName;
                //if (elementName.Contains("管束") || elementName.Contains("束帶") || elementName.Contains("吊架"))
                if (targetPara != null && targetPara.AsString().Contains("CEC-吊架"))
                {
                    if (element.Category.Name == "管附件")
                    {
                        {
                            return true;
                        }
                    }
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }
    }
}