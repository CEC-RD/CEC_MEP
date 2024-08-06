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
using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace AutoHangerCreation_ButtonCreate
{
    //調整吊架長度&更新吊架系統別
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ThreadAdjust : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Centimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //針對吊架元件進行碰撞檢討測試
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            #region 舊作法，以吊架為基礎去撞管
            ////製作品類過濾器
            //BuiltInCategory[] builts =
            //{
            //BuiltInCategory.OST_PipeCurves,
            //BuiltInCategory.OST_Conduit,
            //BuiltInCategory.OST_DuctCurves
            //};
            //List<ElementFilter> filters = new List<ElementFilter>();
            //foreach (BuiltInCategory built in builts)
            //{
            //    ElementCategoryFilter filter = new ElementCategoryFilter(built);
            //    filters.Add(filter);
            //}
            //LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
            //foreach (ElementId id in ids)
            //{
            //    Element elem = doc.GetElement(id);
            //    FamilyInstance inst = elem as FamilyInstance;
            //    FilteredElementCollector pipeCollector = new FilteredElementCollector(doc).WherePasses(categoryFilter);
            //    BoundingBoxXYZ castBounding = inst.get_BoundingBox(null);
            //    Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
            //    BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);
            //    Solid castSolid = method.singleSolidFromElement(inst);
            //    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
            //    pipeCollector.WherePasses(boxIntersectsFilter).WherePasses(solidFilter);
            //    List<Element> pipeCollector_final = new List<Element>();
            //    foreach (Element e in pipeCollector)
            //    {
            //        pipeCollector_final.Add(e);
            //    }
            //    MessageBox.Show($"我是id為{inst.Id}的吊架，與我關的管共有{pipeCollector_final.Count()}個");
            //}
            #endregion

            #region 新作法->用管去撞吊架
            foreach (ElementId id in ids)
            {
                Element elem = doc.GetElement(id);
                BuiltInCategory[] builts =
                {
                BuiltInCategory.OST_PipeAccessory,
                };
                List<ElementFilter> filters = new List<ElementFilter>();
                foreach (BuiltInCategory built in builts)
                {
                    ElementCategoryFilter filter = new ElementCategoryFilter(built);
                    filters.Add(filter);
                }

                List<FamilyInstance> targeInstLst = new List<FamilyInstance>();
                List<ElementId> idList = new List<ElementId>();
                LogicalOrFilter categoryFilter = new LogicalOrFilter(filters);
                FilteredElementCollector hangerCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WherePasses(categoryFilter);
                foreach (Element e in hangerCollector)
                {
                    if (findTargetElement(e))
                    {
                        FamilyInstance inst = e as FamilyInstance;
                        targeInstLst.Add(inst);
                        idList.Add(inst.Id);
                    }
                }
                MessageBox.Show(idList.Count().ToString());
                FilteredElementCollector collector = new FilteredElementCollector(doc, idList);
                BoundingBoxXYZ castBounding = elem.get_BoundingBox(null);
                Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);
                Solid castSolid = method.singleSolidFromElement(elem);
                ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                collector.WherePasses(boxIntersectsFilter).WherePasses(solidFilter);
                MessageBox.Show($"我是id為{elem.Id}的管，與我干涉的吊架共有{collector.Count()}個");
            }
            #endregion

            return Result.Succeeded;
        }

        public bool findTargetElement(Element e)
        {
            bool isFound = false;
            List<string> checkName = new List<string>()
            {
                "吊架","多管"
            };
            FamilyInstance inst = e as FamilyInstance;
            FamilySymbol tempSymbol = inst.Symbol;
            Parameter nameToCheck = tempSymbol.LookupParameter("API識別名稱");
            foreach (string st in checkName)
            {
                if (nameToCheck != null && nameToCheck.AsString().Contains(st))
                {
                    isFound = true;
                    break;
                }
            }
            return isFound;
        }
    }
}
