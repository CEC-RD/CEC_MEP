using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI.Selection;
using PipeTagger.Windows;

namespace PipeTagger
{
    [Transaction(TransactionMode.Manual)]
    public class NotFinish : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                NothingWindow window = new NothingWindow(uidoc);
                //window.Label_Name.Content = Properties.Settings.Default.Test_text;
                window.ShowDialog();
                return Result.Succeeded;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                message = e.Message;
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class Tag_setting : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            try
            {
                setting_tag window = new setting_tag(uidoc);
                window.ShowDialog();

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                message = e.Message;
                return Result.Failed;
            }
        }
    }


    [Transaction(TransactionMode.Manual)]
    public class Place_tag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            return Common.Placer.CreateSomething(uidoc, doc, ref message, "Tag");

        }
    }

    [Transaction(TransactionMode.Manual)]
    public class Place_spot : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            return Common.Placer.CreateSomething(uidoc, doc, ref message, "Spot");

        }
    }

}

#region References

#region A. How to add tag
#region 1. "HorizontalContiTagged" -- > Mainv4_1 -->  Line 156 ()

//Go To  "HorizontalContiTagged" -- > Mainv4_1 -->  Line 156 ()


///// <summary>
///// 建立水管管徑標註
///// </summary>
//[Transaction(TransactionMode.Manual)]
//public class CreatPipeDiameterTag : IExternalCommand
//{
//    #region IExternalCommand Members

//    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//    {
//        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
//        Document doc = uiDoc.Document;//當前活動文件
//        Autodesk.Revit.DB.View view = uiDoc.ActiveView;//當前活動檢視
//        Selection sel = uiDoc.Selection;//選擇集
//        Transaction ts = new Transaction(doc, "水管管徑標記");
//        try
//        {
//            ts.Start();
//            PipeSelectionFilter psf = new PipeSelectionFilter(doc);
//            Reference refer = sel.PickObject(ObjectType.Element, psf, "請選擇要標註的水管:");
//            Pipe pipe = doc.GetElement(refer) as Pipe;
//            if (pipe == null)
//            {
//                ts.Dispose();
//                TaskDialog.Show("RevitMassge", "沒有選中水管");
//                return Result.Failed;
//            }
//            //Define tag mode and tag orientation for new tag
//            TagMode tageMode = TagMode.TM_ADDBY_CATEGORY;
//            TagOrientation tagOri = TagOrientation.Horizontal;
//            //Add the tag to the middle of duct
//            LocationCurve locCurve = pipe.Location as LocationCurve;
//            XYZ pipeMid = locCurve.Curve.Evaluate(0.275, true);


//            IndependentTag tag = IndependentTag.Create(doc, 399884, view.Id, ); //(https://www.revitapidocs.com/2019/b8e8eec2-8e3b-08f2-a9a5-89f24465c8b9.htm)
//                                                                                //view, pipe, false, tageMode, tagOri, pipeMid)';

//            //遍歷型別
//            FilteredElementCollector filterColl = GetElementsOfType(doc, typeof(FamilySymbol), BuiltInCategory.OST_PipeTags);
//            //WinFormTools.MsgBox(filterColl.ToElements().Count.ToString());
//            int elId = 0;
//            foreach (Element el in filterColl.ToElements())
//            {
//                if (el.Name == "管道尺寸標記")
//                    elId = el.Id.IntegerValue;
//            }
//            tag.ChangeTypeId(new ElementId(elId));
//            ElementId eId = null;
//            if (tag == null)
//            {
//                ts.Dispose();
//                TaskDialog.Show("RevitMassge", "建立標註失敗!");
//                return Result.Failed;
//            }
//            ICollection<ElementId> eSet = tag.GetValidTypes();
//            foreach (ElementId item in eSet)
//            {
//                if (item.IntegerValue == 532753)
//                {
//                    eId = item;
//                }
//            }
//            tag = doc.GetElement(eId) as IndependentTag;
//        }
//        catch (Exception)
//        {
//            ts.Dispose();
//            return Result.Cancelled;
//        }
//        ts.Commit();

//        return Result.Succeeded;
//    }
//    FilteredElementCollector GetElementsOfType(Document doc, Type type, BuiltInCategory bic)
//    {
//        FilteredElementCollector collector = new FilteredElementCollector(doc);

//        collector.OfCategory(bic);
//        collector.OfClass(type);

//        return collector;
//    }
//    #endregion
//}

///// <summary>
/////水管選擇過濾器
///// </summary>
//public class PipeSelectionFilter : ISelectionFilter
//{
//    #region ISelectionFilter Members

//    Document doc = null;
//    public PipeSelectionFilter(Document document)
//    {
//        doc = document;
//    }

//    public bool AllowElement(Element elem)
//    {
//        return elem is Pipe;
//    }

//    public bool AllowReference(Reference reference, XYZ position)
//    {
//        return doc.GetElement(reference) is Pipe;
//    }

//    #endregion
//}

#endregion

#region 2. Online example  https://forums.autodesk.com/t5/revit-api-forum/independenttag-how-do-i-call-this-in-revit/td-p/7733731

//[Transaction(TransactionMode.Manual)]
//public class Tagtest : IExternalCommand
//{
//    #region Methods

//    /// <summary>
//    ///       The CreateIndependentTag
//    /// </summary>
//    /// <param name="document">The <see cref="Document" /></param>
//    /// <param name="wall">The <see cref="Wall" /></param>
//    /// <returns>The <see cref="IndependentTag" /></returns>
//    public IndependentTag CreateIndependentTag(Document document, Wall wall)
//    {
//        TaskDialog.Show("Create Independent Tag Method", "Start Of Method Dialog");
//        // make sure active view is not a 3D view
//        var view = document.ActiveView;

//        // define tag mode and tag orientation for new tag
//        var tagMode = TagMode.TM_ADDBY_CATEGORY;
//        var tagorn = TagOrientation.Horizontal;

//        // Add the tag to the middle of the wall
//        var wallLoc = wall.Location as LocationCurve;
//        var wallStart = wallLoc.Curve.GetEndPoint(0);
//        var wallEnd = wallLoc.Curve.GetEndPoint(1);
//        var wallMid = wallLoc.Curve.Evaluate(0.5, true);
//        var wallRef = new Reference(wall);

//        var newTag = IndependentTag.Create(document, view.Id, wallRef, true, tagMode, tagorn, wallMid);
//        if (null == newTag) throw new Exception("Create IndependentTag Failed.");

//        // newTag.TagText is read-only, so we change the Type Mark type parameter to 
//        // set the tag text.  The label parameter for the tag family determines
//        // what type parameter is used for the tag text.

//        var type = wall.WallType;

//        var foundParameter = type.LookupParameter("Type Mark");
//        //var result = foundParameter.Set("Hello");

//        // set leader mode free
//        // otherwise leader end point move with elbow point

//        newTag.LeaderEndCondition = LeaderEndCondition.Free;
//        var elbowPnt = wallMid + new XYZ(5.0, 5.0, 0.0);
//        newTag.LeaderElbow = elbowPnt;
//        var headerPnt = wallMid + new XYZ(10.0, 10.0, 0.0);
//        newTag.TagHeadPosition = headerPnt;

//        TaskDialog.Show("Create Independent Tag Method", "End Of Method Dialog");

//        return newTag;
//    }

//    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//    {
//        UIDocument uidoc = commandData.Application.ActiveUIDocument;
//        Document doc = uidoc.Document;

//        try
//        {
//            Reference r = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);

//            Wall w = doc.GetElement(r.ElementId) as Wall;

//            using (Transaction t = new Transaction(doc, "create tag"))
//            {
//                t.Start();

//                CreateIndependentTag(doc, w);

//                t.Commit();
//            }
//        }
//        catch (Exception e)
//        {
//            message = e.Message;
//            return Result.Failed;
//        }

//        return Result.Succeeded;
//    }

//    #endregion
//}

#endregion

#region 3. Add tag for Linked reference https://forums.autodesk.com/t5/revit-api-forum/tagging-linked-elements-using-revit-api/m-p/8671548#M37534
/// <summary>
/// Tag all walls in all linked documents
/// </summary>
//void TagAllLinkedWalls(Document doc)
//{
//    // Point near my wall
//    XYZ xyz = new XYZ(-20, 20, 0);

//    // At first need to find our links
//    FilteredElementCollector collector
//        = new FilteredElementCollector(doc)
//        .OfClass(typeof(RevitLinkInstance));

//    foreach (Element elem in collector)
//    {
//        // Get linkInstance
//        RevitLinkInstance instance = elem
//            as RevitLinkInstance;

//        // Get linkDocument
//        Document linkDoc = instance.GetLinkDocument();

//        // Get linkType
//        RevitLinkType type = doc.GetElement(
//            instance.GetTypeId()) as RevitLinkType;

//        // Check if link is loaded
//        if (RevitLinkType.IsLoaded(doc, type.Id))
//        {
//            // Find walls for tagging
//            FilteredElementCollector walls
//                = new FilteredElementCollector(linkDoc)
//                .OfCategory(BuiltInCategory.OST_Walls)
//                .OfClass(typeof(Wall));

//            // Create reference
//            foreach (Wall wall in walls)
//            {
//                Reference newRef = new Reference(wall)
//                    .CreateLinkReference(instance);

//                // Create transaction
//                using (Transaction tx = new Transaction(doc))
//                {
//                    tx.Start("Create tags");

//                    IndependentTag newTag = IndependentTag.Create(
//                        doc, doc.ActiveView.Id, newRef, true,
//                        TagMode.TM_ADDBY_MATERIAL,
//                        TagOrientation.Horizontal, xyz);

//                    // Use TaggedElementId.LinkInstanceId and 
//                    // TaggedElementId.LinkInstanceId to retrieve 
//                    // the id of the tagged link and element:

//                    LinkElementId linkId = newTag.TaggedElementId;
//                    ElementId linkInsancetId = linkId.LinkInstanceId;
//                    ElementId linkedElementId = linkId.LinkedElementId;

//                    tx.Commit();
//                }
//            }
//        }
//    }
//}


#endregion

#endregion

#region How to change tag style


#region 1. Online example https://adndevblog.typepad.com/aec/2013/01/setting-the-leader-arrowhead-for-a-structural-framing-tag-using-revit-api.html

//namespace Revit.SDK.Samples.HelloRevit.CS

//{

//    [Transaction(TransactionMode.Manual)]
//    public class Command : IExternalCommand{
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication uiApp = commandData.Application;
//            Document doc = uiApp.ActiveUIDocument.Document;

//            // Access all elements in the model which represent Arrowheads
//            // This is being done by filtering all elements which
//            // are of ElementType and have the ALL_MODEL_FAMILY_NAME
//            // builtIn Parameter set to 'Arrowhead'

//            ElementId id = new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME);
//            ParameterValueProvider provider = new ParameterValueProvider(id);
//            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
//            FilterRule rule = new FilterStringRule(provider, evaluator, "Arrowhead", false);
//            ElementParameterFilter filter = new ElementParameterFilter(rule);
//            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ElementType)).WherePasses(filter);


//            using (Transaction trans = new Transaction(doc, "Arrowhead"))
//            {
//                trans.Start();
//                // For simplicity, assuming that the
//                // Structural Component Tag is selected
//                foreach (Element selectedElement in
//                  uiApp.ActiveUIDocument.Selection.Elements)
//                {
//                    IndependentTag tag = selectedElement as IndependentTag;

//                    if (null != tag)
//                    {
//                        // Access the Symbol of the IndependentTag element
//                        FamilySymbol tagSymbol = doc.GetElement(tag.GetTypeId()) as FamilySymbol;
//                        // Set the LEADER_ARROWHEAD parameter of the
//                        // Symbol with one of the arrowheads that was filtered
//                        tagSymbol.get_Parameter(BuiltInParameter.LEADER_ARROWHEAD).Set(collector.ToElementIds().ElementAt<ElementId>(0));
//                    }
//                    trans.Commit();
//                }
//            }
//            return Result.Succeeded;
//        }
//    }
//}

#endregion

#endregion

#endregion