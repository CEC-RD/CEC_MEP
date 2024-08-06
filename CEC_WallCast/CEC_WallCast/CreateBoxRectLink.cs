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
    class CreateBoxRectLink : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //準備東西
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //拿到外參電器設備
            ISelectionFilter linkBoxFilter = new linkedEqipSelectionFilter(doc);
            Reference refer = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkBoxFilter);
            RevitLinkInstance boxLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
            Transform boxLinkedTrans = boxLinkedInst.GetTotalTransform();
            Element linkedBox = boxLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);

            //拿到整份外參牆&Tranform
            ISelectionFilter linkedWallFilter = new WallSelectionFilter(doc);
            Reference pickWallRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedWallFilter, "請選擇「連結模型」中被貫穿的牆");
            RevitLinkInstance pickWall = doc.GetElement(pickWallRef) as RevitLinkInstance;
            Transform linkTransform = pickWall.GetTotalTransform();

            Element linkedWall = pickWall.GetLinkDocument().GetElement(pickWallRef.LinkedElementId);
            Wall wall = linkedWall as Wall;
            double holeLength = UnitUtils.ConvertFromInternalUnits(wall.Width, unitType) + 20;
            LocationCurve wallLocate = linkedWall.Location as LocationCurve;
            Curve wallCrv = wallLocate.Curve;
            wallCrv = wallCrv.CreateTransformed(linkTransform);
            Line wallLine = wallCrv as Line;
            double angle = 0.0;
            bool isLiner = false;
            XYZ holeDir = XYZ.BasisY;
            if(wallLine != null)
            {
                isLiner = true;
                XYZ wallDir = wallLine.Direction.Normalize();
                XYZ wallNorDir = wallDir.CrossProduct(XYZ.BasisZ).Normalize().Negate(); //這段要再看一下
                angle = holeDir.AngleTo(wallNorDir);
            }

            //Solid boxSolid = 


            return Result.Succeeded;
        }
    }

    public class linkedEqipSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public linkedEqipSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            return true;
        }
        public bool AllowReference(Reference refer, XYZ point)
        {
            //Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            //Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            //Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            //Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            Category electricalEqip = Category.GetCategory(_doc, BuiltInCategory.OST_ElectricalEquipment);
            var elem = this._doc.GetElement(refer);
            if (elem != null && elem is RevitLinkInstance link)
            {
                var linkElem = link.GetLinkDocument().GetElement(refer.LinkedElementId);
                if (linkElem.Category.Id == electricalEqip.Id)
                {
                    return true;
                }
            }
            else
            {
                if (elem.Category.Id == electricalEqip.Id)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
