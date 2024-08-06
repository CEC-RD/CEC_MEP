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

#endregion

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AddHangerByMouseLink : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Centimeters;
#endif
        public static string errorOutput = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            while (true)
            {
                Counter.count += 1;
                try
                {
                    UIDocument uidoc = commandData.Application.ActiveUIDocument;
                    Document doc = uidoc.Document;
                    Autodesk.Revit.UI.Selection.ISelectionFilter linkPipeFilter = new linkedPipeSelectionFilter(doc);
                    XYZ previousPt = null;
                    Reference previousRefer = null;
                    ReferenceArray reArray = null;
                    //PreviewControl pControl = new PreviewControl(doc, doc.ActiveView.Id);
                    //pControl.MouseMove += MessageBox.Show("");
                    Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, linkPipeFilter, "請點選欲放置管架的位置");
                    XYZ position = refer.GlobalPoint;
                    bool isLinked = true;
                    //取得管的locationCurve後進行投影，取得垂直於管上的點位
                    //Element elem = doc.GetElement(refer.ElementId);
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    if (pipeLinkedInst != null)
                    {
                        isLinked = true;
                    }
                    else
                    {
                        isLinked = false;
                        Element e = doc.GetElement(refer.ElementId);
                    }
                    Curve curve = null;
                    Element elem = null;
                    Line pipeLine = null;
                    if (isLinked)
                    {
                        Transform pipeLinkedTrans = pipeLinkedInst.GetTotalTransform();
                        elem = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                        //MessageBox.Show(elem.Name);
                        LocationCurve pipeCrv = elem.Location as LocationCurve;
                        curve = pipeCrv.Curve;
                        XYZ startPoint = curve.GetEndPoint(0);
                        XYZ endPoint = curve.GetEndPoint(1);
                        startPoint = method.TransformPoint(startPoint, pipeLinkedTrans);
                        endPoint = method.TransformPoint(endPoint, pipeLinkedTrans);
                        pipeLine = Line.CreateBound(startPoint, endPoint);
                    }
                    else if (!isLinked)
                    {
                        elem = doc.GetElement(refer.ElementId);
                        LocationCurve pipeCrv = elem.Location as LocationCurve;
                        curve = pipeCrv.Curve;
                        pipeLine = curve as Line;

                    }

                    IntersectionResult intersect = curve.Project(position);
                    XYZ targetPoint = intersect.XYZPoint;
                    //根據選中的管徑進行元件類型的篩選
                    Parameter targetPara = null;
                    switch (elem.Category.Name)
                    {
                        case "管":
                            targetPara = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                            break;
                        case "電管":
                            targetPara = elem.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                            break;
                        case "風管":
                            targetPara = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                            break;
                    }
                    if (targetPara == null)
                    {
                        MessageBox.Show("目前還暫不適用方形管件，請待後續更新");
                    }
                    using (TransactionGroup transactionGroup = new TransactionGroup(doc))
                    {
                        transactionGroup.Start("放置單管吊架");
                        using (Transaction trans = new Transaction(doc))
                        {

                            trans.Start("放置單管吊架");
                            var degrees = 0.0;
                            double half_PI = Math.PI / 2;
                            double pipeDia = targetPara.AsDouble();
                            XYZ target_up = new XYZ(targetPoint.X, targetPoint.Y, targetPoint.Z + 1);

                            XYZ rotateBase = new XYZ(0, targetPoint.X, 0);
                            Line Axis = Line.CreateBound(targetPoint, target_up);

                            FamilySymbol targetSymbol = new pipeHanger().getFamilySymbol(doc, pipeDia);
                            Element hanger = new pipeHanger().CreateHanger(uidoc.Document, targetPoint, elem, targetSymbol);

                            degrees = rotateBase.AngleTo(pipeLine.Direction);
                            double a = degrees * 180 / (Math.PI);
                            double finalRotate = Math.Abs(half_PI - degrees);
                            if (a > 135 || a < 45)
                            {
                                finalRotate = -finalRotate;
                            }
                            //旋轉後校正位置

                            hanger.Location.Rotate(Axis, finalRotate);
                            previousPt = targetPoint;
                            previousRefer = new Reference(hanger);
                            FamilyInstance instance = hanger as FamilyInstance;
                            double adjust = UnitUtils.ConvertToInternalUnits(1000, unitType);
                            XYZ upperPoint = new XYZ(targetPoint.X, targetPoint.Y, targetPoint.Z + adjust);
                            if (method.checkPara(instance, "PipeCenterToFloor"))
                            {
                                instance.LookupParameter("PipeCenterToFloor").Set(adjust);
                            }
                            else
                            {
                                MessageBox.Show("請確認元件中是否有「PipeCenterToFloor」參數");
                            }
                            FamilyInstance inst = hanger as FamilyInstance;
                            method.updateHangerContent(inst, new List<Element>() { elem });
                            trans.Commit();


                            trans.Start("吊架與版資訊蒐集");
                            Dictionary<ElementId, List<Element>> hangerSlabDict = method.getHangerSlabDict(doc, instance);

                            Line intersectCrv = Line.CreateBound(targetPoint, upperPoint);

                            List<XYZ> intersectPts = new List<XYZ>();

                            if (hangerSlabDict[instance.Id].Count() == 0)
                            {
                                return Result.Failed;
                            }
                            foreach (Element linkSlab in hangerSlabDict[instance.Id])
                            {

                                RevitLinkInstance targetLink = method.getTargetLinkedInstance(doc, linkSlab.Document.Title);
                                Transform linkedInstTrans = targetLink.GetTotalTransform();
                                Solid slabSolid = method.singleSolidFromElement(linkSlab);
                                slabSolid = SolidUtils.CreateTransformed(slabSolid, linkedInstTrans);
                                SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();

                                if (null != slabSolid)
                                {
                                    SolidCurveIntersection intersection = slabSolid.IntersectWithCurve(intersectCrv, options);
                                    int intersectCount = intersection.SegmentCount;
                                    if (intersectCount > 0)
                                    {
                                        Curve castIntersect_Crv = intersection.GetCurveSegment(0);
                                        XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                                        XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);
                                        intersectPts.Add(intersect_DN);
                                    }
                                }
                            }
                            intersectPts = intersectPts.OrderBy(x => x.Z).ToList();
                            double finalLength = intersectPts.First().Z - targetPoint.Z;
                            string threadLenName = "PipeCenterToFloor";
                            instance.LookupParameter(threadLenName).Set(finalLength);
                            trans.Commit();
                        }
                        transactionGroup.Assimilate();
                    }
                }
                catch
                {
                    break;
                }
            }
            Counter.count += 1;
            return Result.Succeeded;
        }
    }
}
