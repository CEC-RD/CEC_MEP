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

namespace BeamCasing_ButtonCreate
{
    class Method
    {
        private XYZ TransformPoint(XYZ point, Transform transform)
        {
            double x = point.X;
            double y = point.Y;
            double z = point.Z;

            //transform basis of the old coordinate system in the new coordinate // system
            XYZ b0 = transform.get_Basis(0);
            XYZ b1 = transform.get_Basis(1);
            XYZ b2 = transform.get_Basis(2);
            XYZ origin = transform.Origin;

            //transform the origin of the old coordinate system in the new 
            //coordinate system
            double xTemp = x * b0.X + y * b1.X + z * b2.X + origin.X;
            double yTemp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y;
            double zTemp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z;

            return new XYZ(xTemp, yTemp, zTemp);
        }
    }
    public class BeamOpening
    {
        public bool isSC;
        public bool isSRC;
        public string docName;
        public RevitLinkInstance rcLinkInstance;
        public RevitLinkInstance scLinkInstance;
        public ElementId beamId;
        public List<Element> scBeamsList;
        public List<Element> rcBeamsList;
        public double castWidth;
    }
    public class BeamCast
    {
        //將穿樑套管設置為class，需要符合下列幾種功能
        //1.先匯入我們目前所有的穿樑套管
        //2.再來判斷選中的管徑與是否有穿過梁，以及穿過的樑種類
        //3.如果有則利用穿過的部分為終點，創造穿樑套管與輸入長度

#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif

        //載入RC穿樑套管元件
        public Family BeamCastSymbol(Document doc)
        {
            //尋找RC樑開口.rfa
            string internalNameRC = "CEC-穿樑套管-圓";
            //string RC_CastName = "穿樑套管共用參數_通用模型";
            Family RC_CastType = null;
            ElementFilter RC_CastCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory);
            ElementFilter RC_CastSymbolFilter = new ElementClassFilter(typeof(FamilySymbol));

            LogicalAndFilter andFilter = new LogicalAndFilter(RC_CastCategoryFilter, RC_CastSymbolFilter);
            FilteredElementCollector RC_CastSymbol = new FilteredElementCollector(doc);
            RC_CastSymbol = RC_CastSymbol.WherePasses(andFilter);//這地方有點怪，無法使用andFilter RC_CastSymbolFilter
            bool symbolFound = false;
            foreach (FamilySymbol e in RC_CastSymbol)
            {
                Parameter p = e.LookupParameter("API識別名稱");
                if (p != null && p.AsString().Contains(internalNameRC))
                {
                    symbolFound = true;
                    RC_CastType = e.Family;
                    break;
                }
            }
            if (!symbolFound)
            {
                MessageBox.Show("尚未載入指定的穿樑套管元件!");
            }
            #region 自己載入元件的方法
            ////如果沒有找到，則自己加載
            //if (!symbolFound)
            //{
            //    string filePath = @"D:\Dropbox (CHC Group)\工作人生\組內專案\04.元件製作\穿樑套管\穿樑套管共用參數_通用模型.rfa";
            //    Family family;
            //    bool loadSuccess = doc.LoadFamily(filePath, out family);
            //    if (loadSuccess)
            //    {
            //        RC_CastType = family;
            //    }
            //}
            #endregion
            return RC_CastType;
        }

        //根據不同的管徑，選擇不同的穿樑套管大小
        public FamilySymbol findRC_CastSymbol(Document doc, Family CastFamily, Element element)
        {
            FamilySymbol targetFamilySymbol = null; //用來找目標familySymbol
            Parameter targetPara = null;
            //Pipe
            if (element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            }
            //Conduit
            else if (element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
            }
            //Duct
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            }
            else if (element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null)
            {
                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            }
            //利用管徑(doubleType)來判斷
            var covertUnit = UnitUtils.ConvertFromInternalUnits(targetPara.AsDouble(), unitType);
            if (CastFamily != null)
            {
                foreach (ElementId castId in CastFamily.GetFamilySymbolIds())
                {
                    FamilySymbol tempSymbol = doc.GetElement(castId) as FamilySymbol;
                    if (covertUnit >= 50 && covertUnit < 65)
                    {
                        if (tempSymbol.Name == "80mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                    else if (covertUnit >= 65 && covertUnit < 75)
                    {
                        if (tempSymbol.Name == "100mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                    //多出關於電管的判斷
                    else if (covertUnit >= 75 && covertUnit <= 95)
                    {
                        if (tempSymbol.Name == "125mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                    else if (covertUnit >= 100 && covertUnit < 125)
                    {
                        if (tempSymbol.Name == "150mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                    else if (covertUnit >= 125 && covertUnit <= 150)
                    {
                        if (tempSymbol.Name == "150mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                    else
                    {
                        if (tempSymbol.Name == "80mm")
                        {
                            targetFamilySymbol = tempSymbol;
                        }
                    }
                }
            }
            if (targetFamilySymbol == null) MessageBox.Show("請確認穿牆套管元件中是否有對應管徑之族群類型");
            targetFamilySymbol.Activate();
            return targetFamilySymbol;
        }
    }
}
