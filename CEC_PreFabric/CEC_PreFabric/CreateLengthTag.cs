#region Namespaces
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using System.IO;
using System.Threading;
#endregion

namespace CEC_PreFabric
{
    //【程式架構】
    //1.待使用者清除標籤後
    //2.以視圖為單位進行篩選
    //3.針對不同的品類上標籤

    [Transaction(TransactionMode.Manual)]
    public class CreateLengthTag : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            ElementFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_PipeTags);
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfClass(typeof(Family));
            List<Family> pipeTags = new List<Family>();
            List<Family> ductTags = new List<Family>();
            List<Family> conduitTags = new List<Family>();
            List<Family> trayTags = new List<Family>();
            foreach (Element e in coll)
            {
                Family tempFam = e as Family;
                if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_PipeTags)
                {
                    pipeTags.Add(tempFam);
                }
                else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_DuctTags)
                {
                    ductTags.Add(tempFam);
                }
                else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_ConduitTags)
                {
                    conduitTags.Add(tempFam);
                }
                else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_CableTrayTags)
                {
                    trayTags.Add(tempFam);
                }
            }
            MessageBox.Show($"管標籤族群有：{pipeTags.Count()}個\n風管標籤族群有：{ductTags.Count()}個\n電管標籤族群有：{conduitTags.Count()}個\n電纜價標籤族群有：{trayTags.Count()}個");

            //以視圖進行篩選，針對品類選擇要上的管標籤
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            if (activeView.ViewType != ViewType.ThreeD)
            {
                MessageBox.Show("本視圖並非3D視圖，請至3D視圖使用此功能");
                return Result.Failed;
            }
            try
            {
                //製作多重品類篩選器
                BuiltInCategory[] builts = {
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_CableTray
                };
                ElementFilter categoriesFilter = multiCateFilter(builts);
                FilteredElementCollector collector = new FilteredElementCollector(doc, activeView.Id).WherePasses(categoriesFilter);
                using (TransactionGroup transGroup = new TransactionGroup(doc))
                {
                    transGroup.Start();
                    foreach (Element e in collector)
                    {
                        placeTag(e);
                    }
                    transGroup.Assimilate();
                }

                MessageBox.Show(collector.Count().ToString());
            }
            catch
            {
                MessageBox.Show("執行失敗");
            }

            return Result.Succeeded;
        }
        public ElementFilter multiCateFilter(BuiltInCategory[] builts)
        {
            List<ElementFilter> filters = new List<ElementFilter>();
            foreach (BuiltInCategory built in builts)
            {
                ElementCategoryFilter filter = new ElementCategoryFilter(built);
                filters.Add(filter);
            }
            LogicalOrFilter categoriesFilter = new LogicalOrFilter(filters);
            return categoriesFilter;
        }
        public void placeTag(Element mepElement)
        {
            Document doc = mepElement.Document;
            Category tempCategory = mepElement.Category;
            BuiltInCategory builtCate = (BuiltInCategory)tempCategory.Id.IntegerValue;

            //蒐集所有需要用到的標籤
            List<string> targetNameList = new List<string>() { "L_管裁切長度標籤", "L_風管裁切長度標籤", "L_電管裁切長度", "L_電纜架裁切標籤" };
            List<Family> targetTags = new List<Family>();
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfClass(typeof(Family));
            foreach(Element e in coll)
            {
                Family tempFam = e as Family;
                if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_PipeTags&&targetNameList.Contains(tempFam.Name))
                {
                    targetTags.Add(tempFam);
                }
                else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_DuctTags && targetNameList.Contains(tempFam.Name))
                {
                    targetTags.Add(tempFam);
                }
                #region 針對電管與風管接可使用
                //else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_ConduitTags && targetNameList.Contains(tempFam.Name))
                //{
                //    targetTags.Add(tempFam);
                //}
                //else if ((BuiltInCategory)tempFam.FamilyCategory.Id.IntegerValue == BuiltInCategory.OST_CableTrayTags && targetNameList.Contains(tempFam.Name))
                //{
                //    targetTags.Add(tempFam);
                //}
                #endregion
            }

            //針對要放的元件品類選擇相對應的標籤元件
            Autodesk.Revit.DB.View isoView = doc.ActiveView;
            FamilySymbol symbolTag=null;
            foreach(Family fam in targetTags)
            {
                if(builtCate == BuiltInCategory.OST_PipeCurves)
                {
                    if (fam.Name == targetNameList[0]) symbolTag = doc.GetElement(fam.GetFamilySymbolIds().First()) as FamilySymbol;
                }
                else if(builtCate == BuiltInCategory.OST_DuctCurves)
                {
                    if (fam.Name == targetNameList[1]) symbolTag = doc.GetElement(fam.GetFamilySymbolIds().First()) as FamilySymbol;
                }
                #region 針對電管與風管接可使用
                //else if (builtCate == BuiltInCategory.OST_Conduit)
                //{
                //    if (fam.Name == targetNameList[2]) symbolTag = doc.GetElement(fam.GetFamilySymbolIds().First()) as FamilySymbol;
                //}
                //else if (builtCate == BuiltInCategory.OST_CableTray)
                //{
                //    if (fam.Name == targetNameList[3]) symbolTag = doc.GetElement(fam.GetFamilySymbolIds().First()) as FamilySymbol;
                //}
                #endregion
            }

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("裁切長度標籤放置");
                Reference elemRefer = new Reference(mepElement);
                MEPCurve mepCrv = mepElement as MEPCurve;
                LocationCurve pipeLocate = mepCrv.Location as LocationCurve;
                Curve pipeCrv = pipeLocate.Curve;
                XYZ middlePt = pipeCrv.Evaluate(0.5, true);
                if (symbolTag != null)
                {
                    IndependentTag fabricTag = IndependentTag.Create(doc, symbolTag.Id, isoView.Id, elemRefer, false, TagOrientation.Horizontal, middlePt);
                }
                trans.Commit();
            }

        }
    }
}
