#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.IO;
#endregion

namespace BeamCasing_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CastInfromUpdateV3 : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw.Reset();//碼表歸零
            sw.Start();//碼表開始計時
            try
            {
                //程式要做四件事如下:
                //1.抓到所有的外參樑
                //2.抓到所有被實做出來的套管
                //3.讓套管和外參樑做交集，比較回傳的參數值是否一樣，不一樣則調整
                //4.並把這些套管亮顯或ID寫下來

                //步驟上進行如下，先挑出所有的RC樑，檢查RC樑中是否有ST樑，
                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                //一些初始設定，套管名稱，檔案名稱，系統縮寫
                string CastMakerName = "CEC";
                string InternalName_ST = "CEC-穿樑開口";
                string InternalName_RC = "CEC-穿樑套管";
                List<string> AllLinkName = new List<string>();
                List<string> systemName = new List<string>() { "E", "T", "W", "P", "F", "A", "G" };
                List<string> usefulParaName = new List<string> { "BTOP", "BCOP", "BBOP", "TTOP", "TCOP", "TBOP", "【原則檢討】上部檢討", "【原則檢討】下部檢討", "【原則檢討】尺寸檢討", "【原則檢討】是否穿樑", "【原則檢討】邊距檢討", "干涉管數量", "系統別", "貫穿樑尺寸", "貫穿樑材料", "貫穿樑編號" };

                //先設定兩個變數判斷是否為SRC樑，如果是的話，基準點也會不一樣
                Solid targetSolid = null;
                Element targetBeam = null;
                Document linkedSC_file = null;
                Document linkedRC_file = null;

                //製作一個容器放所有被實做出來的套管元件，先篩出所有doc中的familyInstance
                //再把指定名字的實體元素加入容器中
                List<FamilyInstance> castInstances = new List<FamilyInstance>();
                FilteredElementCollector coll = new FilteredElementCollector(doc);
                ElementCategoryFilter castCate_Filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
                ElementClassFilter castInst_Filter = new ElementClassFilter(typeof(Instance));
                LogicalAndFilter andFilter = new LogicalAndFilter(castCate_Filter, castInst_Filter);
                coll.WherePasses(andFilter).WhereElementIsNotElementType().ToElements(); //找出模型中實做的穿樑套管元件
                if (coll != null)
                {
                    foreach (FamilyInstance e in coll)
                    {
                        //if (e.Symbol.FamilyName == CastName)
                        //以製造商跟API識別名稱雙重確認是否為正確的元件
                        //if (e.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).AsString() == CastMakerName && e.Symbol.FamilyName.Contains("套管"))
                        if (e.Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).AsString() == CastMakerName && e.Symbol.LookupParameter("API識別名稱").AsString().Contains(InternalName_ST))
                        {
                            castInstances.Add(e);
                        }
                    }
                }

                //檢查蒐集到的元件是否都有相對應的名稱
                foreach (FamilyInstance famInst in castInstances)
                {
                    foreach (string item in usefulParaName)
                    {
                        if (!CheckPara(famInst, item))
                        {
                            MessageBox.Show($"執行失敗，請檢查{famInst.Symbol.FamilyName}元件中是否缺少{item}參數欄位");
                            return Result.Failed;
                        }
                    }
                }

                if (castInstances.Count() == 0)
                {
                    MessageBox.Show("尚未匯入套管元件，或模型中沒有實做的套管元件");
                    return Result.Failed;
                }
                //建置一個List來裝外參中所有的RC樑
                List<Element> RC_Beams = new List<Element>();
                List<Element> ST_Beams = new List<Element>();
                ICollection<ElementId> RC_BeamsID = new List<ElementId>();
                ICollection<ElementId> ST_BeamsID = new List<ElementId>();

                //建置蒐集特例的List
                int updateCastNum = 0;
                List<Element> intersectInst = new List<Element>();
                List<ElementId> updateCastIDs = new List<ElementId>();
                List<ElementId> Cast_tooClose = new List<ElementId>(); //存放離樑頂或樑底太近的套管
                List<ElementId> Cast_tooBig = new List<ElementId>(); //存放太大的套管
                List<ElementId> Cast_Conflict = new List<ElementId>(); //存放彼此太過靠近的套管
                List<ElementId> Cast_BeamConfilct = new List<ElementId>(); //存放大樑兩端過近的套管
                List<ElementId> Cast_OtherConfilct = new List<ElementId>(); //存放小樑兩端過近的套管
                List<ElementId> Cast_Empty = new List<ElementId>();//存放空管的穿樑套管

                ElementCategoryFilter linkedFileFilter = new ElementCategoryFilter(BuiltInCategory.OST_RvtLinks);
                FilteredElementCollector linkedFileCollector = new FilteredElementCollector(doc).WherePasses(linkedFileFilter).WhereElementIsNotElementType();

                //找到是否為SRC的方法
                //先找到鋼構的外參檔案名稱(前提是鋼構和混凝土要完全切開)
                if (linkedFileCollector.Count() > 0)
                {
                    foreach (RevitLinkInstance linkedInst in linkedFileCollector)
                    {
                        Document linkDoc = linkedInst.GetLinkDocument();
                        FilteredElementCollector linkedBeams = new FilteredElementCollector(linkDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                        foreach (Element e in linkedBeams)
                        {
                            FamilyInstance beamInstance = e as FamilyInstance;
                            if (beamInstance.StructuralMaterialType.ToString() == "Steel")
                            {
                                linkedSC_file = linkDoc;
                            }
                        }
                    }
                }
                //再來找到混凝土的外參檔案名稱(前提是鋼構和混凝土要完全切開)
                if (linkedFileCollector.Count() > 0)
                {
                    foreach (RevitLinkInstance linkedInst in linkedFileCollector)
                    {
                        Document linkDoc = linkedInst.GetLinkDocument();
                        FilteredElementCollector linkedBeams = new FilteredElementCollector(linkDoc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                        foreach (Element e in linkedBeams)
                        {
                            FamilyInstance beamInstance = e as FamilyInstance;
                            if (beamInstance.StructuralMaterialType.ToString() == "Concrete")
                            {
                                linkedRC_file = linkDoc;
                            }
                        }
                    }
                }

                List<Element> All_RCBeams = new List<Element>();
                if (null != linkedRC_file)
                {
                    FilteredElementCollector RC_collector = new FilteredElementCollector(linkedRC_file).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                    All_RCBeams = RC_collector.WhereElementIsNotElementType().ToList();
                }

                //如果模型中只有RC檔，則SC和RC應該是同一個
                if (null == linkedSC_file)
                {
                    linkedSC_file = linkedRC_file;
                }
                string mistakeReportRC = "";
                string mistakeReportSC = "";
                string testGrider = "";
                if (null != linkedSC_file)
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("更新穿樑套管資訊測試");
                        foreach (Element linkedElement in All_RCBeams)
                        {
                            //檢查每隻樑是否為SRC與大樑
                            bool isSRC = false;
                            bool isGrider = false;
                            FilteredElementCollector SRC_collector = new FilteredElementCollector(linkedSC_file).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                            Solid RCSolid = singleSolidFromElement(linkedElement);
                            if (RCSolid == null)
                            {
                                mistakeReportRC += $"我是編號{linkedElement.Id}的樑，我來自{linkedElement.Document.Title}，我無法創造一個完整的實體\n";
                                continue;
                            }
                            BoundingBoxXYZ RCBounding = linkedElement.get_BoundingBox(null);
                            Autodesk.Revit.DB.Transform t = RCBounding.Transform;
                            Outline outline = new Outline(t.OfPoint(RCBounding.Min), t.OfPoint(RCBounding.Max));
                            //Outline outline = new Outline(RCBounding.Min, RCBounding.Max);
                            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outline);
                            ElementIntersectsSolidFilter elementIntersectsSolidFilter = new ElementIntersectsSolidFilter(RCSolid);
                            SRC_collector.WherePasses(boundingBoxIntersectsFilter).WherePasses(elementIntersectsSolidFilter);
                            if (SRC_collector.Count() > 0)
                            {
                                isSRC = true;
                            }
                            ////先取得在這根樑中所有的穿樑套管 (效能會比較差)
                            //FilteredElementCollector tempCollector = new FilteredElementCollector(doc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_PipeAccessory);
                            //tempCollector.WherePasses(boundingBoxIntersectsFilter).WherePasses(elementIntersectsSolidFilter);
                            //List<Element> castInThisBeam = new List<Element>();
                            //foreach (Element e in tempCollector)
                            //{
                            //    string instFamName = e.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                            //    if (instFamName.Contains("穿樑套管共用參數_通用模型"))
                            //    {
                            //        castInThisBeam.Add(e);
                            //    }
                            //}

                            //if (castInThisBeam.Count() > 0)
                            //{
                            //    MessageBox.Show($"這根樑中有{castInThisBeam.Count()}個套管");
                            //}
                            //if (SRC_collector.Count() > 0 /*&& castInThisBeam.Count() > 0*/)
                            if (isSRC)
                            {
                                Solid SCSolid = singleSolidFromElement(SRC_collector.First());
                                if (SRC_collector.First() != null && SCSolid != null)
                                {
                                    targetSolid = SCSolid;
                                    targetBeam = SRC_collector.First();
                                }
                                else
                                {
                                    mistakeReportSC += $"我是編號{SRC_collector.First().Id}的樑，我來自{SRC_collector.First().Document.Title}，我無法創造一個完整的實體\n";
                                    continue;
                                }
                            }

                            else if (!isSRC)
                            {
                                targetSolid = RCSolid;
                                targetBeam = linkedElement;
                                if (targetSolid == null) continue;
                            }

                            if (null != targetSolid)
                            {
                                SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();
                                foreach (FamilyInstance inst in castInstances)
                                {
                                    LocationPoint instLocate = inst.Location as LocationPoint;
                                    double inst_CenterZ = (inst.get_BoundingBox(null).Max.Z + inst.get_BoundingBox(null).Min.Z) / 2;
                                    XYZ instPt = instLocate.Point;
                                    double normal_BeamHeight = UnitUtils.ConvertToInternalUnits(1000, unitType);
                                    XYZ inst_Up = new XYZ(instPt.X, instPt.Y, instPt.Z + normal_BeamHeight);
                                    XYZ inst_Dn = new XYZ(instPt.X, instPt.Y, instPt.Z - normal_BeamHeight);
                                    Curve instVerticalCrv = Autodesk.Revit.DB.Line.CreateBound(inst_Dn, inst_Up);
                                    //這邊用solid是因為怕有斜樑需要開口的問題，但斜樑的結構應力應該已經蠻集中的，不可以再開口
                                    SolidCurveIntersection intersection = targetSolid.IntersectWithCurve(instVerticalCrv, options);
                                    int intersectCount = intersection.SegmentCount;
                                    //針對有切割到的實體去做計算
                                    if (intersectCount > 0)
                                    {
                                        string instInternalName = inst.Symbol.LookupParameter("API識別名稱").AsString();
                                        //針對有交集的實體去做計算
                                        inst.LookupParameter("【原則檢討】是否穿樑").Set("OK");
                                        //intersectInst.Add(inst);

                                        //計算TOP、BOP等六個參數
                                        LocationPoint cast_Locate = inst.Location as LocationPoint;
                                        XYZ LocationPt = cast_Locate.Point;
                                        XYZ cast_Max = inst.get_BoundingBox(null).Max;
                                        XYZ cast_Min = inst.get_BoundingBox(null).Min;
                                        Curve castIntersect_Crv = intersection.GetCurveSegment(0);
                                        XYZ intersect_DN = castIntersect_Crv.GetEndPoint(0);
                                        XYZ intersect_UP = castIntersect_Crv.GetEndPoint(1);
                                        double castCenter_Z = (cast_Max.Z + cast_Min.Z) / 2;

                                        double TTOP_update = intersect_UP.Z - cast_Max.Z;
                                        double BTOP_update = cast_Max.Z - intersect_DN.Z;
                                        double TCOP_update = intersect_UP.Z - castCenter_Z;
                                        double BCOP_update = castCenter_Z - intersect_DN.Z;
                                        double TBOP_update = intersect_UP.Z - cast_Min.Z;
                                        double BBOP_update = cast_Min.Z - intersect_DN.Z;
                                        double TTOP_orgin = inst.LookupParameter("TTOP").AsDouble();
                                        double BBOP_orgin = inst.LookupParameter("BBOP").AsDouble();
                                        double beamHeight = intersect_UP.Z - intersect_DN.Z;
                                        double castHeight = cast_Max.Z - cast_Min.Z;
                                        //計算套管的高度，注意有Symbol和沒有Symbol的取法
                                        //double castHeight = 0.0;
                                        //if (instInternalName.Contains("圓"))
                                        //{
                                        //    castHeight = inst.Symbol.LookupParameter("管外直徑").AsDouble();
                                        //}else if (instInternalName.Contains("方"))
                                        //{
                                        //    castHeight = inst.LookupParameter("高度(H)").AsDouble();
                                        //}
                                        double TTOP_Check = Math.Round(UnitUtils.ConvertFromInternalUnits(TTOP_update, unitType), 1);
                                        double TTOP_orginCheck = Math.Round(UnitUtils.ConvertFromInternalUnits(TTOP_orgin, unitType), 1);
                                        double BBOP_Check = Math.Round(UnitUtils.ConvertFromInternalUnits(BBOP_update, unitType), 1);
                                        double BBOP_orginCheck = Math.Round(UnitUtils.ConvertFromInternalUnits(BBOP_orgin, unitType), 1);


                                        //檢查樑為大樑還是小樑，更新檢核的判斷依據
                                        double C_distRatio = 0.0;
                                        double C_protectRatio = 0.0;
                                        double C_protectMin = 0.0;
                                        double C_sizeRatio = 0.0;
                                        double C_sizeMax = 0.0;
                                        double R_distRatio = 0.0;
                                        double R_protectRatio = 0.0;
                                        double R_protectMin = 0.0;
                                        double R_sizeRatioD = 0.0;
                                        double R_sizeRatioW = 0.0;

                                        //確認是大樑還小樑後，重新賦值
                                        isGrider = CheckGrider(targetBeam);
                                        if (isGrider)
                                        {
                                            C_distRatio = BeamCast_Settings.Default.cD1_Ratio;
                                            C_protectRatio = BeamCast_Settings.Default.cP1_Ratio;
                                            C_protectMin = BeamCast_Settings.Default.cP1_Min;
                                            C_sizeRatio = BeamCast_Settings.Default.cMax1_Ratio;
                                            C_sizeMax = BeamCast_Settings.Default.cMax1_Max;
                                            R_distRatio = BeamCast_Settings.Default.rD1_Ratio;
                                            R_protectRatio = BeamCast_Settings.Default.rP1_Ratio;
                                            R_protectMin = BeamCast_Settings.Default.rP1_Min;
                                            R_sizeRatioD = BeamCast_Settings.Default.rMax1_RatioD;
                                            R_sizeRatioW = BeamCast_Settings.Default.rMax1_RatioW;
                                            testGrider += $"編號{targetBeam.Id}的樑是大樑\n";
                                        }
                                        else if (!isGrider)
                                        {
                                            C_distRatio = BeamCast_Settings.Default.cD2_Ratio;
                                            C_protectRatio = BeamCast_Settings.Default.cP2_Ratio;
                                            C_protectMin = BeamCast_Settings.Default.cP2_Min;
                                            C_sizeRatio = BeamCast_Settings.Default.cMax2_Ratio;
                                            C_sizeMax = BeamCast_Settings.Default.cMax2_Max;
                                            R_distRatio = BeamCast_Settings.Default.rD2_Ratio;
                                            R_protectRatio = BeamCast_Settings.Default.rP2_Ratio;
                                            R_protectMin = BeamCast_Settings.Default.rP2_Min;
                                            R_sizeRatioD = BeamCast_Settings.Default.rMax2_RatioD;
                                            R_sizeRatioW = BeamCast_Settings.Default.rMax2_RatioW;
                                            testGrider += $"編號{targetBeam.Id}的樑是小樑\n";
                                        }

                                        List<double> parameter_Checklist = new List<double> { C_distRatio, C_protectRatio, C_sizeRatio, R_distRatio, R_protectRatio, R_sizeRatioD, R_sizeRatioW };
                                        List<double> parameter_Checklist2 = new List<double> { C_protectMin, C_sizeMax, R_protectMin };

                                        //string test = "";
                                        //foreach (double d in parameter_Checklist)
                                        //{
                                        //    test+= $"{d}\n";
                                        //}
                                        //MessageBox.Show(test);

                                        foreach (double d in parameter_Checklist)
                                        {
                                            if (d == 0)
                                            {
                                                message = "穿樑原則的設定值有誤(非選填數值不可為0)，請重新設定後再次執行";
                                                return Result.Failed;
                                            }
                                        }
                                        for (int i = 0; i < parameter_Checklist2.Count; i++)
                                        {
                                            if (parameter_Checklist2[i] == null)
                                            {
                                                parameter_Checklist2[i] = 0;
                                            }
                                        }

                                        //只要六個數值中有任何一個<0，代表沒有穿樑成功
                                        List<double> updateParas = new List<double> { TTOP_update, BTOP_update, TCOP_update, BCOP_update, TBOP_update, BBOP_update };
                                        foreach (double d in updateParas)
                                        {
                                            if (d < 0)
                                            {
                                                inst.LookupParameter("【原則檢討】是否穿樑").Set("不符合");
                                            }
                                        }

                                        if (TTOP_Check != TTOP_orginCheck || BBOP_Check !=BBOP_orginCheck)
                                        {
                                            inst.LookupParameter("TTOP").Set(TTOP_update);
                                            inst.LookupParameter("BTOP").Set(BTOP_update);
                                            inst.LookupParameter("TCOP").Set(TCOP_update);
                                            inst.LookupParameter("BCOP").Set(BCOP_update);
                                            inst.LookupParameter("TBOP").Set(TBOP_update);
                                            inst.LookupParameter("BBOP").Set(BBOP_update);
                                            Element updateElem = inst as Element;
                                        }

                                        //寫入樑編號與樑尺寸
                                        string beamName = targetBeam.LookupParameter("編號").AsString();
                                        string beamSIze = targetBeam.LookupParameter("類型").AsValueString();//抓取類型
                                        if (beamName != null)
                                        {
                                            inst.LookupParameter("貫穿樑編號").Set(beamName);
                                        }
                                        else
                                        {
                                            inst.LookupParameter("貫穿樑編號").Set("無編號");
                                        }
                                        if (beamSIze != null)
                                        {
                                            inst.LookupParameter("貫穿樑尺寸").Set(beamSIze);
                                        }
                                        else
                                        {
                                            inst.LookupParameter("貫穿樑尺寸").Set("無尺寸");
                                        }


                                        //太過靠近樑底的套管
                                        double alertValue = 0.0; //設定樑底與樑頂的距離警告
                                        double tempProtectValue = 0.0;
                                        if (instInternalName.Contains("圓"))
                                        {
                                            alertValue = C_protectRatio * beamHeight;
                                            tempProtectValue = UnitUtils.ConvertToInternalUnits(C_protectMin, unitType);
                                        }
                                        else if (instInternalName.Contains("方"))
                                        {
                                            alertValue = R_protectRatio * beamHeight;
                                            tempProtectValue = UnitUtils.ConvertToInternalUnits(R_protectMin, unitType);
                                        }
                                        if (tempProtectValue > alertValue)
                                        {
                                            alertValue = tempProtectValue;
                                        }
                                        if (TTOP_update < alertValue || BBOP_update < alertValue)
                                        {
                                            Cast_tooClose.Add(inst.Id);
                                            if (TTOP_update < alertValue)
                                            {
                                                inst.LookupParameter("【原則檢討】上部檢討").Set("不符合");
                                            }
                                            else if (BBOP_update < alertValue)
                                            {
                                                inst.LookupParameter("【原則檢討】下部檢討").Set("不符合");
                                            }
                                        }
                                        else
                                        {
                                            inst.LookupParameter("【原則檢討】上部檢討").Set("OK");
                                            inst.LookupParameter("【原則檢討】下部檢討").Set("OK");
                                        }

                                        //太大的圓套管
                                        if (instInternalName.Contains("圓"))
                                        {
                                            double alertMaxSize = C_sizeRatio * beamHeight;//設定最大尺寸警告
                                            double tempSizeValue = UnitUtils.ConvertToInternalUnits(C_sizeMax, unitType);
                                            if (tempSizeValue < alertValue)
                                            {
                                                alertMaxSize = tempSizeValue;
                                            }
                                            if (castHeight > alertMaxSize)
                                            {
                                                Cast_tooBig.Add(inst.Id);
                                                inst.LookupParameter("【原則檢討】尺寸檢討").Set("不符合");
                                            }
                                            else
                                            {
                                                inst.LookupParameter("【原則檢討】尺寸檢討").Set("OK");
                                            }
                                        }
                                        else if (instInternalName.Contains("方"))
                                        {
                                            double alertMaxSizeD = R_sizeRatioD * beamHeight;//設定最大高度尺寸檢討
                                            double alertMaxSizeW = R_sizeRatioW * beamHeight;//設定最大寬度尺寸檢討
                                            double instHeight = inst.LookupParameter("高度(H)").AsDouble();
                                            double instWidth = inst.LookupParameter("寬度(W)").AsDouble();
                                            if (instHeight > alertMaxSizeD || instWidth > alertMaxSizeW)
                                            {
                                                inst.LookupParameter("【原則檢討】尺寸檢討").Set("不符合");
                                            }
                                            else
                                            {
                                                inst.LookupParameter("【原則檢討】尺寸檢討").Set("OK");
                                            }
                                        }


                                        //距離大樑(StructuralType=Beam) 或小樑(StructuralType=Other) 太近的套管
                                        //先判斷是方開口還是圓開口(距離不一樣)
                                        FamilyInstance beamInst = (FamilyInstance)targetBeam;
                                        LocationCurve tempLocateCrv = null;
                                        Curve targetCrv = null;
                                        XYZ startPt = null;
                                        XYZ endPt = null;
                                        List<XYZ> points = new List<XYZ>();
                                        string BeamUsage = beamInst.StructuralUsage.ToString();
                                        string BeamMaterial = beamInst.StructuralMaterialType.ToString();

                                        if (BeamMaterial == "Concrete")
                                        {
                                            inst.LookupParameter("貫穿樑材料").Set("RC開孔");
                                            tempLocateCrv = beamInst.Location as LocationCurve;
                                            targetCrv = tempLocateCrv.Curve;
                                            XYZ tempStart = targetCrv.GetEndPoint(0);
                                            XYZ tempEnd = targetCrv.GetEndPoint(1);
                                            startPt = new XYZ(tempStart.X, tempStart.Y, instPt.Z);
                                            endPt = new XYZ(tempEnd.X, tempStart.Y, instPt.Z);
                                            points.Add(startPt);
                                            points.Add(endPt);
                                            foreach (XYZ pt in points)
                                            {
                                                double distToBeamEnd = instPt.DistanceTo(pt);
                                                //判斷是方形還是圓型，調整參數
                                                double distCheck = 0.0;
                                                if (instInternalName.Contains("圓"))
                                                {
                                                    distCheck = C_distRatio * beamHeight;
                                                }
                                                else if (instInternalName.Contains("方"))
                                                {
                                                    distCheck = R_distRatio * beamHeight;
                                                }

                                                if (distToBeamEnd < distCheck)
                                                {
                                                    Cast_BeamConfilct.Add(inst.Id);
                                                    inst.LookupParameter("【原則檢討】邊距檢討").Set("不符合");
                                                }
                                                else
                                                {
                                                    inst.LookupParameter("【原則檢討】邊距檢討").Set("OK");
                                                }
                                            }
                                        }
                                        else if (BeamMaterial == "Steel")
                                        {
                                            inst.LookupParameter("貫穿樑材料").Set("鋼構開孔");
                                            tempLocateCrv = beamInst.Location as LocationCurve;
                                            targetCrv = tempLocateCrv.Curve;
                                            XYZ tempStart = targetCrv.GetEndPoint(0);
                                            XYZ tempEnd = targetCrv.GetEndPoint(1);
                                            startPt = new XYZ(tempStart.X, tempStart.Y, instPt.Z);
                                            endPt = new XYZ(tempEnd.X, tempStart.Y, instPt.Z);
                                            points.Add(startPt);
                                            points.Add(endPt);
                                            foreach (XYZ pt in points)
                                            {
                                                double distToBeamEnd = instPt.DistanceTo(pt);
                                                //判斷是方形還是圓型，調整參數
                                                double distCheck = 0.0;
                                                if (instInternalName.Contains("圓"))
                                                {
                                                    distCheck = C_distRatio * beamHeight;
                                                }
                                                else if (instInternalName.Contains("方"))
                                                {
                                                    distCheck = R_distRatio * beamHeight;
                                                }
                                                if (distToBeamEnd - castHeight / 2 < beamHeight)
                                                {
                                                    Cast_OtherConfilct.Add(inst.Id);
                                                    inst.LookupParameter("【原則檢討】邊距檢討").Set("不符合");
                                                    break;
                                                }
                                                else
                                                {
                                                    inst.LookupParameter("【原則檢討】邊距檢討").Set("OK");
                                                }
                                            }
                                        }

                                        //離別人太近的套管
                                        //(在原本的list中去除自己，逐個量測距離，取出彼此靠太近的套管，去除重複ID後，由大到小排序
                                        int index = castInstances.IndexOf(inst);
                                        for (int i = 0; i < castInstances.Count(); i++)
                                        {
                                            if (i == index)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                //利用嚴謹的方法求取距離
                                                //如果是圓形的套管，會有三種狀況，圓與圓、方與方、圓與方
                                                //圓開口和圓開口測距離
                                                if (instInternalName.Contains("圓") && castInstances[i].Symbol.LookupParameter("API識別名稱").AsString().Contains("圓"))
                                                {
                                                    double Dia1 = inst.Symbol.LookupParameter("管外直徑").AsDouble();
                                                    double Dia2 = castInstances[i].Symbol.LookupParameter("管外直徑").AsDouble();
                                                    LocationPoint thisLocation = inst.Location as LocationPoint;
                                                    LocationPoint otherLocation = castInstances[i].Location as LocationPoint;
                                                    XYZ thisPt = thisLocation.Point;
                                                    XYZ otherPt = otherLocation.Point;
                                                    XYZ newPt = new XYZ(otherPt.X, otherPt.Y, thisPt.Z);
                                                    double distBetween = thisPt.DistanceTo(newPt);
                                                    if (distBetween < (Dia1 + Dia2) * 1.5)
                                                    {
                                                        Cast_Conflict.Add(castInstances[i].Id);
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("不符合");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("OK");
                                                    }
                                                }
                                                //圓開口和方開口測距離
                                                else if (instInternalName.Contains("圓") && castInstances[i].Symbol.LookupParameter("API識別名稱").AsString().Contains("方"))
                                                {
                                                    double Dia1 = inst.Symbol.LookupParameter("管外直徑").AsDouble();
                                                    double Dia2 = castInstances[i].LookupParameter("寬度(W)").AsDouble();
                                                    LocationPoint thisLocation = inst.Location as LocationPoint;
                                                    LocationPoint otherLocation = castInstances[i].Location as LocationPoint;
                                                    XYZ thisPt = thisLocation.Point;
                                                    XYZ otherPt = otherLocation.Point;
                                                    XYZ newPt = new XYZ(otherPt.X, otherPt.Y, thisPt.Z);
                                                    double distBetween = thisPt.DistanceTo(newPt);
                                                    if (distBetween < (Dia1 + Dia2) * 1.5 || distBetween / 1.5 < beamHeight)
                                                    {
                                                        Cast_Conflict.Add(castInstances[i].Id);
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("不符合");
                                                         break;
                                                    }
                                                    else
                                                    {
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("OK");
                                                    }
                                                }
                                                //方開口和方開口測距離
                                                else if (instInternalName.Contains("方") && castInstances[i].Symbol.LookupParameter("API識別名稱").AsString().Contains("方"))
                                                {
                                                    double Dia1 = inst.LookupParameter("寬度(W)").AsDouble();
                                                    double Dia2 = castInstances[i].LookupParameter("寬度(W)").AsDouble();
                                                    LocationPoint thisLocation = inst.Location as LocationPoint;
                                                    LocationPoint otherLocation = castInstances[i].Location as LocationPoint;
                                                    XYZ thisPt = thisLocation.Point;
                                                    XYZ otherPt = otherLocation.Point;
                                                    XYZ newPt = new XYZ(otherPt.X, otherPt.Y, thisPt.Z);
                                                    double distBetween = thisPt.DistanceTo(newPt);
                                                    if (distBetween / 1.5 < beamHeight)
                                                    {
                                                        Cast_Conflict.Add(castInstances[i].Id);
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("不符合");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        castInstances[i].LookupParameter("【原則檢討】邊距檢討").Set("OK");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //針對套管去做更新：貫穿管的系統別、貫穿的管隻數
                        using (SubTransaction subTrans = new SubTransaction(doc))
                        {
                            subTrans.Start();
                            foreach (FamilyInstance inst in castInstances)
                            {
                                FilteredElementCollector pipeCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves);
                                BoundingBoxXYZ castBounding = inst.get_BoundingBox(null);
                                Outline castOutline = new Outline(castBounding.Min, castBounding.Max);
                                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(castOutline);
                                Solid castSolid = singleSolidFromElement(inst);
                                ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(castSolid);
                                pipeCollector.WherePasses(boxIntersectsFilter).WherePasses(solidFilter);
                                inst.LookupParameter("干涉管數量").Set(pipeCollector.Count());

                                if (pipeCollector.Count() == 0)
                                {
                                    Cast_Empty.Add(inst.Id);
                                    inst.LookupParameter("系統別").Set("未指定");
                                }
                                //針對蒐集到的管去做系統別的更新
                                if (pipeCollector.Count() == 1)
                                {
                                    string pipeSystem = pipeCollector.First().get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                                    string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                                    if (systemName.Contains(shortSystemName))
                                    {
                                        inst.LookupParameter("系統別").Set(shortSystemName);
                                    }
                                    else
                                    {
                                        inst.LookupParameter("系統別").Set("未指定");
                                    }
                                }
                                //如果有共管的狀況
                                else if (pipeCollector.Count() >= 2)
                                {
                                    List<string> shortNameList = new List<string>();
                                    foreach (Element pipe in pipeCollector)
                                    {
                                        string pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                                        string shortSystemName = pipeSystem.Substring(0, 1);//以開頭的前綴字抓系統縮寫
                                        shortNameList.Add(shortSystemName);
                                        List<string> newList = shortNameList.Distinct().ToList();
                                        //就算共管，如果同系統還是得寫一樣的縮寫名稱
                                        if (newList.Count() == 1)
                                        {
                                            inst.LookupParameter("系統別").Set(newList.First());
                                        }
                                        //如果為不同系統共管，則設為M
                                        else if (newList.Count() > 1)
                                        {
                                            inst.LookupParameter("系統別").Set("M");
                                        }
                                    }
                                }
                            }
                            subTrans.Commit();
                        }

                        string path = @"C:\Users\12061753\Desktop\Test.txt";
                        File.WriteAllText(path, mistakeReportRC);
                        File.AppendAllText(path, mistakeReportSC);
                        File.AppendAllText(path, testGrider);
                        sw.Stop();//碼錶停止
                        double sec = sw.Elapsed.TotalMilliseconds / 1000;
                        string result1 = sec.ToString();
                        MessageBox.Show($"穿樑套管更新完成，共花費 {sec} 秒");
                        trans.Commit();
                    }



                    //將UI視窗實做出來
                    CastInformUpdateUI updateWindow = new CastInformUpdateUI(commandData);
                    updateWindow.Show();
                    //設定Source源
                    if (Cast_tooClose.Count > 0)
                    {
                        updateWindow.ProtectConflictListBox.ItemsSource = Cast_tooClose;
                    }
                    else
                    {
                        updateWindow.ProtectConflictListBox.ItemsSource = "無";
                    }
                    if (Cast_Conflict.Count > 0)
                    {
                        updateWindow.TooCloseCastListBox.ItemsSource = Cast_Conflict;
                    }
                    else
                    {
                        updateWindow.TooCloseCastListBox.ItemsSource = "無";
                    }

                    if (Cast_tooBig.Count > 0)
                    {
                        updateWindow.TooBigCastListBox.ItemsSource = Cast_tooBig;
                    }
                    else
                    {
                        updateWindow.TooBigCastListBox.ItemsSource = "無";
                    }

                    if (Cast_OtherConfilct.Count > 0)
                    {
                        updateWindow.OtherCastListBox.ItemsSource = Cast_OtherConfilct;
                    }
                    else
                    {
                        updateWindow.OtherCastListBox.ItemsSource = "無";
                    }

                    if (Cast_BeamConfilct.Count > 0)
                    {
                        updateWindow.GriderCastListBox.ItemsSource = Cast_BeamConfilct;
                    }
                    else
                    {
                        updateWindow.GriderCastListBox.ItemsSource = "無";
                    }

                    if (Cast_Empty.Count > 0)
                    {
                        updateWindow.EmptyCastListBox.ItemsSource = Cast_Empty;
                    }
                    else
                    {
                        updateWindow.EmptyCastListBox.ItemsSource = "無";
                    }
                }
            }
            catch
            {
                //如果執行失敗，先檢查元件名稱，以及是否有BTOP、BCOP、BBOP、TTOP、TCOP、TBOP等六個參數
                //在進一步檢查元件名稱，以及是否有 【原則檢討】上部檢討、【原則檢討】下部檢討、【原則檢討】尺寸檢討、【原則檢討】邊距檢討、【原則檢討】是否穿樑 等六個參數
                //進一步檢查是否有干涉管數量、系統別、貫穿樑尺寸、貫穿樑編號
                message = "執行失敗，請檢查元件版次與參數是否遭到更改!";
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        //製做一個方法，取的樑的solid
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
            // create solid from Element:
            IList<Solid> fromElement = GetTargetSolids(inputElement);
            int solidCount = fromElement.Count;
            // MessageBox.Show(solidCount.ToString());
            // Merge all found solids into single one
            Solid solidResult = null;
            //XYZ checkheight = new XYZ(0, 0, 6.88976);
            //Transform tr = Transform.CreateTranslation(checkheight);
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

        public bool CheckPara(Element elem, string paraName)
        {
            bool result = false;
            foreach (Parameter parameter in elem.Parameters)
            {
                Parameter val = parameter;
                if (val.Definition.Name == paraName)
                {
                    result = true;
                }
            }
            return result;
        }

        public bool CheckGrider(Element elem)
        {
            bool result = false;
            Document doc = elem.Document;
            FilteredElementCollector tempCollector = new FilteredElementCollector(doc).OfClass(typeof(Instance)).OfCategory(BuiltInCategory.OST_StructuralColumns);
            BoundingBoxXYZ checkBounding = elem.get_BoundingBox(null);
            Autodesk.Revit.DB.Transform t1 = checkBounding.Transform;
            Outline outline1 = new Outline(t1.OfPoint(checkBounding.Min), t1.OfPoint(checkBounding.Max));
            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter1 = new BoundingBoxIntersectsFilter(outline1, 0.1);
            //ElementIntersectsSolidFilter elementIntersectsSolidFilter1 = new ElementIntersectsSolidFilter(RCSolid);
            tempCollector.WherePasses(boundingBoxIntersectsFilter1);
            if (tempCollector.Count() > 0)
            {
                result = true;
            }
            else if (tempCollector.Count() == 0)
            {
                result = false;
            }
            return result;
        }

        //public List<Element> getAllLinkBeams (Document )
    }
}
