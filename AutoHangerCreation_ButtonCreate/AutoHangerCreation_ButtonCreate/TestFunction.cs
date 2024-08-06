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
using BuildingCoder;

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TestFunction : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Centimeters;
#endif
        //製作一個功能測試多管吊架和樓版之間的接觸關係
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Element e = doc.GetElement(ids.First());
            BoundingBoxXYZ hangerBox = e.get_BoundingBox(null);
            MessageBox.Show($"吊架boundingBox的最大值為{hangerBox.Max}，最小值為{hangerBox.Min}");
            List<Element> conflictSlabs = new List<Element>();
            List<RevitLinkInstance> slabLinks = findLinkInstWithSlab(doc);
            foreach(RevitLinkInstance linkInst in slabLinks)
            {
                Transform trans = linkInst.GetTotalTransform();
                Transform inversetrans = trans.Inverse;
                Outline outLine = new Outline(inversetrans.OfPoint(hangerBox.Min), inversetrans.OfPoint(hangerBox.Max));
                BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outLine);
                FilteredElementCollector collectorSB = getAllLinkedSlab(linkInst.GetLinkDocument());
                collectorSB.WherePasses(boundingBoxIntersectsFilter);
                MessageBox.Show($"與編號{e.Id}吊架碰撞的樓板共有{collectorSB.Count()}個");
            }
            return Result.Succeeded;
        }
        public FilteredElementCollector getAllLinkedSlab(Document linkedDoc)
        {
            FilteredElementCollector slabCollector = new FilteredElementCollector(linkedDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            return slabCollector;
        }
        public List<RevitLinkInstance> findLinkInstWithSlab(Document doc)
        {
            List<RevitLinkInstance> targetLinkedInstances = new List<RevitLinkInstance>();
            ElementCategoryFilter linkedFileFilter = new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks);
            FilteredElementCollector linkedFileCollector = new FilteredElementCollector(doc).WherePasses(linkedFileFilter).WhereElementIsNotElementType();
            try
            {
                if (linkedFileCollector.Count() > 0)
                {
                    foreach (RevitLinkInstance linkedInst in linkedFileCollector)
                    {
                        Document linkDoc = linkedInst.GetLinkDocument();
                        bool isLoaded = RevitLinkType.IsLoaded(doc, linkedInst.GetTypeId());
                        if (linkDoc != null && isLoaded)
                        {
                            FilteredElementCollector linkedSlab = new FilteredElementCollector(linkDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                            if (linkedSlab.Count() == 0) continue;
                            else
                            {
                                MessageBox.Show($"{linkedInst.Name}中有{linkedSlab.Count()}個樓板");
                                if (!targetLinkedInstances.Contains(linkedInst)) targetLinkedInstances.Add(linkedInst);
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("請檢查外參連結是否載入或有問題!");
            }
            return targetLinkedInstances;
        }
        public IList<Solid> GetTargetSolids(Element element)
        {
            List<Solid> solids = new List<Solid>();
            Options options = new Options();
            //預設為不包含不可見元件，因此改成true
            options.ComputeReferences = true;
            options.DetailLevel = ViewDetailLevel.Fine;
            options.IncludeNonVisibleObjects = true;
            GeometryElement geomElem = element.get_Geometry(options);
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid)
                {
                    Solid solid = (Solid)geomObj;
                    if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                    {
                        solids.Add(solid);
                    }
                }
                else if (geomObj is GeometryInstance)//一些特殊狀況可能會用到，like樓梯
                {
                    GeometryInstance geomInst = (GeometryInstance)geomObj;
                    GeometryElement instGeomElem = geomInst.GetInstanceGeometry();
                    foreach (GeometryObject instGeomObj in instGeomElem)
                    {
                        if (instGeomObj is Solid)
                        {
                            Solid solid = (Solid)instGeomObj;
                            if (solid.Faces.Size > 0 && solid.Volume > 0.0)
                            {
                                solids.Add(solid);
                            }
                        }
                    }
                }
            }
            return solids;
        }
        public Solid singleSolidFromElement(Element inputElement)
        {
            Document doc = inputElement.Document;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            IList<Solid> fromElement = GetTargetSolids(inputElement);
            int solidCount = fromElement.Count;
            Solid solidResult = null;
            if (solidCount == 1)
            {
                solidResult = fromElement[0];
            }
            else if (solidCount > 1)
            {
                solidResult =
                    BooleanOperationsUtils.ExecuteBooleanOperation(fromElement[0], fromElement[1], BooleanOperationsType.Union);
            }

            if (solidCount > 2)
            {
                for (int i = 2; i < solidCount; i++)
                {
                    solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(solidResult, fromElement[i], BooleanOperationsType.Union);
                }
            }
            return solidResult;
        }
    }
}
