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
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Threading;
#endregion

namespace CEC_Count
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class EquipCount : IExternalCommand
    {
        private Thread _uiThread;
        // ModelessForm instance
        private CountingUI _mMyForm;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            this.ShowForm(commandData);
            //this.ShowFormSeparateThread(commandData);
            return Result.Succeeded;
        }
        public DirectShape createSolidFromBBox(Autodesk.Revit.DB.View view)
        {
            Document doc = view.Document;

            BoundingBoxXYZ inputBb = null;
            double cutPlaneHeight = 0.0;
            XYZ pt0 = null;
            XYZ pt1 = null;
            XYZ pt2 = null;
            XYZ pt3 = null;
            Solid preTransformBox = null;
            if (view.ViewType == ViewType.FloorPlan)
            {
                inputBb = view.get_BoundingBox(null);
                Autodesk.Revit.DB.Plane planePlanView = view.SketchPlane.GetPlane();
                Autodesk.Revit.DB.PlanViewRange viewRange = (view as Autodesk.Revit.DB.ViewPlan).GetViewRange();
                cutPlaneHeight = viewRange.GetOffset(Autodesk.Revit.DB.PlanViewPlane.CutPlane);
                //XYZ pt0 = inputBb.Min;
                //XYZ pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                //XYZ pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                //XYZ pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                double level = view.GenLevel.ProjectElevation;
                pt0 = new XYZ(inputBb.Min.X, inputBb.Min.Y, level);
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, level);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, level);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, level);

                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, cutPlaneHeight);
                //Solid 
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
            }
            else if (view.ViewType == ViewType.ThreeD)
            {
                View3D view3D = (View3D)view;
                //inputBb = view3D.GetSectionBox();
                inputBb = view.CropBox;
                if (inputBb == null) MessageBox.Show("請確認剖面框是否開啟");
                pt0 = inputBb.Min;
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, solidheight);
            }
            // Put this inside a transaction!


            DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            ds.ApplicationId = "Test";
            ds.ApplicationDataId = "testBox";
            List<GeometryObject> GeoList = new List<GeometryObject>();
            GeoList.Add(preTransformBox); // <-- the solid created for the intersection can be used here
            ds.SetShape(GeoList);
            ds.SetName("ID_testBox");

            return ds;
        }

        public void ShowForm(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            // If we do not have a dialog yet, create and show it
            Document doc = uiapp.ActiveUIDocument.Document;
            if (_mMyForm != null && _mMyForm == null) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();
            ZoomHandlerWithWpfArg eZoom = new ZoomHandlerWithWpfArg();
            UpdateHandlerWithWpfArg eUpdate = new UpdateHandlerWithWpfArg();
            #region

            // The dialog becomes the owner responsible for disposing the objects given to it.
            #endregion
            _mMyForm = new CountingUI(commandData, evStr, evWpf,eZoom,eUpdate);
            _mMyForm.Show();
        }

        public Solid getSolidFromBBox(Autodesk.Revit.DB.View view)
        {
            Document doc = view.Document;

            BoundingBoxXYZ inputBb = null;
            double cutPlaneHeight = 0.0;
            XYZ pt0 = null;
            XYZ pt1 = null;
            XYZ pt2 = null;
            XYZ pt3 = null;
            Solid preTransformBox = null;
            if (view.ViewType == ViewType.FloorPlan)
            {
                inputBb = view.get_BoundingBox(null);
                Autodesk.Revit.DB.Plane planePlanView = view.SketchPlane.GetPlane();
                Autodesk.Revit.DB.PlanViewRange viewRange = (view as Autodesk.Revit.DB.ViewPlan).GetViewRange();
                cutPlaneHeight = viewRange.GetOffset(Autodesk.Revit.DB.PlanViewPlane.CutPlane);
                //XYZ pt0 = inputBb.Min;
                //XYZ pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                //XYZ pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                //XYZ pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                double level = view.GenLevel.ProjectElevation;
                pt0 = new XYZ(inputBb.Min.X, inputBb.Min.Y, level);
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, level);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, level);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, level);

                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, cutPlaneHeight);
                //Solid 
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
            }
            else if (view.ViewType == ViewType.ThreeD)
            {
                View3D view3D = (View3D)view;
                inputBb = view3D.GetSectionBox();
                if (inputBb == null) MessageBox.Show("請確認剖面框是否開啟");
                pt0 = inputBb.Min;
                pt1 = new XYZ(inputBb.Max.X, inputBb.Min.Y, inputBb.Min.Z);
                pt2 = new XYZ(inputBb.Max.X, inputBb.Max.Y, inputBb.Min.Z);
                pt3 = new XYZ(inputBb.Min.X, inputBb.Max.Y, inputBb.Min.Z);
                Line edge0 = Line.CreateBound(pt0, pt1);
                Line edge1 = Line.CreateBound(pt1, pt2);
                Line edge2 = Line.CreateBound(pt2, pt3);
                Line edge3 = Line.CreateBound(pt3, pt0);
                List<Curve> edges = new List<Curve>();
                edges.Add(edge0);
                edges.Add(edge1);
                edges.Add(edge2);
                edges.Add(edge3);
                CurveLoop baseLoop = CurveLoop.Create(edges);
                List<CurveLoop> loopList = new List<CurveLoop>();
                loopList.Add(baseLoop);
                double solidheight = inputBb.Max.Z - inputBb.Min.Z;
                preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, solidheight);
            }
            return preTransformBox;
        }
        public void ShowFormSeparateThread(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            // If we do not have a thread started or has been terminated start a new one
            if (!(_uiThread is null) && _uiThread.IsAlive) return;
            //EXTERNAL EVENTS WITH ARGUMENTS
            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();
            ZoomHandlerWithWpfArg eZoom = new ZoomHandlerWithWpfArg();
            UpdateHandlerWithWpfArg eUpdate = new UpdateHandlerWithWpfArg();

            //新增執行序
            _uiThread = new Thread(() =>
            {
                SynchronizationContext.SetSynchronizationContext(
                    new DispatcherSynchronizationContext(
                        Dispatcher.CurrentDispatcher));
                // The dialog becomes the owner responsible for disposing the objects given to it.

                _mMyForm = new CountingUI(commandData, evStr, evWpf,eZoom,eUpdate);
                _mMyForm.Closed += (s, e) => Dispatcher.CurrentDispatcher.InvokeShutdown();
                _mMyForm.Show();
                Dispatcher.Run();
            });

            _uiThread.SetApartmentState(ApartmentState.STA);
            _uiThread.IsBackground = true;
            _uiThread.Start();
        }
    }
}
