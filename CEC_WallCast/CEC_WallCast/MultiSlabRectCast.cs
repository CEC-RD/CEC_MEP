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
    class MultiSlabRectCast : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;
                method m = new method();

                //拿到管元件(多管)
                ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                ISelectionFilter linkedPipeFilter = new linkedPipeSelectionFilter(doc);
                IList<Reference> pickPipeRefs = uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, "請選擇「本端模型」中貫穿版的管");
                List<Element> pickPipes = new List<Element>();
                foreach (Reference refer in pickPipeRefs)
                {
                    Element tempPipe = doc.GetElement(refer);
                    pickPipes.Add(tempPipe);
                }
                IList<Reference> linkedPickPipeRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, linkedPipeFilter, "請選擇「連結模型」中貫穿版的管(選填)");
                //string output = "";
                foreach (Reference refer in linkedPickPipeRefs)
                {
                    RevitLinkInstance pipeLinkedInst = doc.GetElement(refer.ElementId) as RevitLinkInstance;
                    Transform pipeLinkedTransform = pipeLinkedInst.GetTotalTransform();
                    Element linkedPipe = pipeLinkedInst.GetLinkDocument().GetElement(refer.LinkedElementId);
                    pickPipes.Add(linkedPipe);
                    //output += $"{linkedPipe.Name}\n";
                }
                //拿到整份牆外參檔&Transform
                ISelectionFilter linkedSlabFilter = new SlabSelectionFilter(doc);
                Reference pickSlabRef = uidoc.Selection.PickObject(ObjectType.LinkedElement, linkedSlabFilter, "請選擇「連結模型」中被貫穿的版");
                RevitLinkInstance pickSlab = doc.GetElement(pickSlabRef) as RevitLinkInstance;
                Transform linkTransform = pickSlab.GetTotalTransform();

                //拿到實際要用的那道版元件
                Document linkDoc = pickSlab.GetLinkDocument();
                Element linkedSlab = linkDoc.GetElement(pickSlabRef.LinkedElementId);
                Level linkLevel = linkDoc.GetElement(linkedSlab.LevelId) as Level;
                double linkLevelElevation = linkLevel.ProjectElevation;
                Level targetLevel = null;
                FilteredElementCollector levelColl = new FilteredElementCollector(doc).OfClass(typeof(Level)).WhereElementIsNotElementType();
                foreach (Element e in levelColl)
                {
                    Level tempLevel = e as Level;
                    if (tempLevel.ProjectElevation == linkLevelElevation)
                    {
                        targetLevel = tempLevel;
                        break;
                    }
                }
                Floor slab = linkedSlab as Floor;
                Solid slabSolid = m.singleSolidFromElement(linkedSlab);
                double slabWidth = linkedSlab.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble();
                double holeLength = UnitUtils.ConvertFromInternalUnits(slabWidth, unitType) + 20;

                //1.擷取每支管和這道牆的交界點
                //2.找出最高的點&最低的點
                //3.找出最左和最右的點
                MEPCurve pipeCrv = pickPipes.First() as MEPCurve;
                LocationCurve pipeLocate = pipeCrv.Location as LocationCurve;
                Curve pipeCurve = pipeLocate.Curve;
                Level level = pipeCrv.ReferenceLevel;//取得管線的參考樓層
                if (level == null)
                {
                    level = doc.GetElement(doc.ActiveView.LevelId) as Level;
                }
                double elevation = level.ProjectElevation;
                List<XYZ> intersectHeight = new List<XYZ>();
                List<XYZ> intersectWidth = new List<XYZ>();
                List<Solid> intersectSolids = new List<Solid>();
                foreach (Element pickPipe in pickPipes)
                {
                    double half_pipeHeight = m.getPipeHeight(pickPipe).AsDouble() / 2;
                    double half_pipeWidth = m.getPipeWidth(pickPipe).AsDouble() / 2;
                    //XYZ widthOffset = half_pipeWidth * wallDir;
                    XYZ holeLocation = m.GetHoleLocation(linkedSlab, pickPipe, linkTransform);
                    if (holeLocation == null)
                    {
                        MessageBox.Show("執行失敗，請確認選中的管是否都有和版交集!");
                        return Result.Failed;
                    }
                    //將pipes和slab進行布林交集
                    Solid tempPipeSolid = m.singleSolidFromElement(pickPipe);
                    Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(slabSolid, tempPipeSolid, BooleanOperationsType.Intersect);
                    if (Math.Abs(interSolid.Volume) > 0.000001)
                    {
                        intersectSolids.Add(interSolid);
                    }
                    //XYZ temp_intersectRightUp = new XYZ(holeLocation.X, holeLocation.Y, holeLocation.Z );
                    //XYZ temp_intersectRightDn = new XYZ(holeLocation.X, holeLocation.Y, holeLocation.Z );
                    //XYZ temp_intersectLeftUp = new XYZ(holeLocation.X , holeLocation.Y, holeLocation.Z);
                    //XYZ temp_intersectLeftDn = new XYZ(holeLocation.X , holeLocation.Y, holeLocation.Z);
                    //intersectHeight.Add(temp_intersectRightUp);
                    //intersectHeight.Add(temp_intersectDn);
                    //intersectWidth.Add(temp_intersectRight);
                    //intersectWidth.Add(temp_intersectLeft);
                }
                #region 根據交集的管進行聚合
                Solid solidResult = null;
                if (intersectSolids.Count == 1)
                {
                    solidResult = intersectSolids[0];
                }
                else if (intersectSolids.Count > 1)
                {
                    solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(intersectSolids[0], intersectSolids[1], BooleanOperationsType.Union);
                }
                if (intersectSolids.Count > 2)
                {
                    for (int i = 2; i < intersectSolids.Count; i++)
                    {
                        solidResult = BooleanOperationsUtils.ExecuteBooleanOperation(solidResult, intersectSolids[i], BooleanOperationsType.Union);
                    }
                }

                //取得要放置穿版套管的正中心點與長寬計算
                //XYZ targetPt = solidResult.ComputeCentroid();
                XYZ targetPt = solidResult.GetBoundingBox().Transform.Origin;
                double holeWidth = solidResult.GetBoundingBox().Max.X - solidResult.GetBoundingBox().Min.X;
                double holeHeight = solidResult.GetBoundingBox().Max.Y - solidResult.GetBoundingBox().Min.Y;
                Family Slab_Cast;
                FamilyInstance instance = null;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("載入檔案測試");
                    Slab_Cast = new RectSlabCast().SlabCastSymbol(doc);
                    tx.Commit();
                }


                #endregion
                using (Transaction trans = new Transaction(doc))
                {
                    trans.Start("放置多穿版套管");
                    //以directShape測試布林結果
                    //DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                    //ds.ApplicationId = "Application id";
                    //ds.ApplicationDataId = "Geometry object id";
                    //ds.SetShape(new GeometryObject[] { solidResult });
                    if (targetPt != null)
                    {
                        FamilySymbol CastSymbol2 = new RectSlabCast().findSlab_CastSymbol(doc, Slab_Cast);
                        instance = doc.Create.NewFamilyInstance(targetPt, CastSymbol2, targetLevel, StructuralType.NonStructural);
                        //參數檢查
                        List<string> paraNameToCheck = new List<string>()
                                {
                                   "開口長","開口高","開口長","系統別"
                                };
                        foreach (string item in paraNameToCheck)
                        {
                            if (!m.checkPara(instance, item))
                            {
                                MessageBox.Show($"執行失敗，請檢查 {instance.Symbol.FamilyName} 元件中是否缺少 {item} 參數欄位");
                                return Result.Failed;
                            }
                        }
#if RELEASE2019
                            Parameter instHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
#else
                        Parameter instHeight = instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM);
#endif

                        double castOffset = UnitUtils.ConvertToInternalUnits(100, unitType);
                        double heightToSet = holeHeight + castOffset;
                        double widthToSet = holeWidth + castOffset;
                        instance.LookupParameter("開口高").Set(heightToSet);
                        instance.LookupParameter("開口寬").Set(widthToSet);
                        instHeight.Set(targetPt.Z - targetLevel.ProjectElevation + slabWidth / 2);
                        double castLength = UnitUtils.ConvertToInternalUnits(holeLength, unitType);
                        instance.LookupParameter("開口長").Set(castLength);
                        trans.Commit();
                    }
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            return Result.Succeeded;
        }

    }
}
