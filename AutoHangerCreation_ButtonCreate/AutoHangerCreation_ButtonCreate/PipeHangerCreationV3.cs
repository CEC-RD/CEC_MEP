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
    //創造單管吊架
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class PipeHangerCreationV3 : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Centimeters;
#endif
        //DisplayUnitType unitType = DisplayUnitType.DUT_CENTIMETERS;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            try
            {
                //限制使用者只能選中管
                UIDocument uidoc = commandData.Application.ActiveUIDocument;

                //點選要一起放置單管吊架的管段
                List<Element> pickElements = new List<Element>();
                Document doc = uidoc.Document;
                Autodesk.Revit.UI.Selection.ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                IList<Reference> pickElements_Refer = uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, $"請選擇欲放置吊架的管段");

                foreach (Reference reference in pickElements_Refer)
                {
                    Element element = doc.GetElement(reference.ElementId);
                    pickElements.Add(element);
                }

                string setupFamily = PIpeHangerSetting.Default.FamilySelected;
                string divideValueString = PIpeHangerSetting.Default.DivideValueSelected;
                if (PIpeHangerSetting.Default.DivideValueSelected == null || PIpeHangerSetting.Default.FamilySelected == null)
                {
                    message = "單管吊架設定未完成!!";
                    return Result.Failed;
                }

                //文字轉數字，英制轉公制
                double divideValue_doubleTemp = double.Parse(divideValueString);
                double divideValue_double = UnitUtils.ConvertToInternalUnits(divideValue_doubleTemp, unitType);
                if (divideValue_double == 0)
                {
                    message = "吊架間距不可為0!!";
                    return Result.Failed;
                }
                //int hangerCount = 0;
                //放置吊架
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("放置單管吊架");
                    foreach (Element element in pickElements)
                    {
                        //先設定一個用來存放目標參數的Para
                        Parameter targetPara = null;
                        switch (element.Category.Name)
                        {
                            case "管":
                                targetPara = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                                break;
                            case "電管":
                                targetPara = element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                                break;
                            case "風管":
                                targetPara = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                                break;
                        }

                        if (targetPara == null)
                        {
                            MessageBox.Show("目前還暫不適用方形管件，請待後續更新");
                            continue;
                        }
                        double pipeDia = targetPara.AsDouble();
                        FamilySymbol targetSymbol = new pipeHanger().getFamilySymbol(doc, pipeDia);

                        //找到管的location curve
                        LocationCurve pipeLocationCrv = element.Location as LocationCurve;
                        Curve pipeCurve = pipeLocationCrv.Curve;

                        XYZ pipeStart = pipeCurve.GetEndPoint(0);
                        XYZ pipeEnd = pipeCurve.GetEndPoint(1);

                        XYZ pipeEndAdjust = new XYZ(pipeEnd.X, pipeEnd.Y, pipeStart.Z);
                        Line pipelineProject = Line.CreateBound(pipeStart, pipeEndAdjust);

                        double pipeLength = pipeCurve.Length;

                        double param1 = pipeCurve.GetEndParameter(0);
                        double param2 = pipeCurve.GetEndParameter(1);

                        //計算要分割的數量 (不確定是否有四捨五入)
                        int step = (int)(pipeLength / divideValue_double);
                        double paramCalc = param1 + ((param2 - param1)
                          * divideValue_double / pipeLength);

                        //創造一個容器裝所有點資料(位於線上的)
                        IList<Point> pointList = new List<Point>();
                        IList<XYZ> locationList = new List<XYZ>();
                        XYZ evaluatedPoint = null;
                        var degrees = 0.0;
                        double half_PI = Math.PI / 2;

                        //如果除出來的階數<=1，則在中點創造一個吊架，如果>1，則依照階數分割
                        if (step <= 1)
                        {
                            XYZ centerPt = pipeCurve.Evaluate(0.5, true);
                            XYZ centerPt_up = new XYZ(centerPt.X, centerPt.Y, centerPt.Z + 1);
                            XYZ rotateBase = new XYZ(0, centerPt.X, 0);
                            Line Axis = Line.CreateBound(centerPt, centerPt_up);
                            Element hanger = new pipeHanger().CreateHanger(uidoc.Document, centerPt, element, targetSymbol);
                            degrees = rotateBase.AngleTo(pipelineProject.Direction);
                            double a = degrees * 180 / (Math.PI);
                            double finalRotate = Math.Abs(half_PI - degrees);
                            if (a > 135 || a < 45)
                            {
                                finalRotate = -finalRotate;
                            }
                            //旋轉後校正位置
                            hanger.Location.Rotate(Axis, finalRotate);
                        }
                        else if (step >= 2)
                        {
                            for (int i = 0; i < step; i++)
                            {
                                paramCalc = param1 + ((param2 - param1) * divideValue_double * (i + 1) / pipeLength);
                                if (pipeCurve.IsInside(paramCalc) == true)
                                {
                                    double normParam = pipeCurve.ComputeNormalizedParameter(paramCalc);
                                    evaluatedPoint = pipeCurve.Evaluate(normParam, true);
                                    Point locationPoint = Point.Create(evaluatedPoint);
                                    pointList.Add(locationPoint);
                                    locationList.Add(evaluatedPoint);
                                }
                            }

                            foreach (XYZ p1 in locationList)
                            {
                                Element hanger = new pipeHanger().CreateHanger(uidoc.Document, p1, element, targetSymbol);
                                XYZ p2 = new XYZ(p1.X, p1.Y, p1.Z + 1);
                                Line Axis = Line.CreateBound(p1, p2);
                                XYZ p3 = new XYZ(0, p1.X, 0); //測量吊架與管段之間的向量差異，取plane中的x向量
                                degrees = p3.AngleTo(pipelineProject.Direction);
                                double a = degrees * 180 / (Math.PI);
                                double finalRotate = Math.Abs(half_PI - degrees);
                                if (a > 135 || a < 45)
                                {
                                    finalRotate = -finalRotate;
                                }
                                //旋轉後校正位置
                                hanger.Location.Rotate(Axis, finalRotate);
                            }
                        }
                    }
                    tx.Commit();
                }
            }
            catch
            {
                //MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            Counter.count += 1;
            //writeDatatoExcel();
            return Result.Succeeded;
        }
        public void writeDatatoExcel()
        {
            #region 紀錄邏輯
            //1.決定要存取的檔案路徑位置&使用者名稱
            //2.判斷檔案是否已經存在該路徑-->沒有則創建檔案，有則打開檔案
            //3.找尋是否有相對應的Column Name(addin Name in English) -->有的話則找到那一整欄，沒有的話則自己創一欄
            //4.找到當下的對應格，以月份去搜尋
            #endregion
            //取得使用者名稱
            //string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string userName = Environment.UserName;
            string fileName = userName + ".xls";
            MessageBox.Show(userName);
            //取得今天的時間
            DateTime toDay = DateTime.Today.Date;
            string theDate = toDay.ToString();
            Excel.Application excelApp = new Excel.Application();
            string filepath = @"C:\Users\12061753\Documents\";
            string finalPath = filepath + fileName;
            if (excelApp != null)
            {
                if (File.Exists(finalPath))
                {
                }
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = new Excel.Worksheet();
                excelWorksheet = excelWorkbook.Worksheets[1];
                excelWorksheet.Name = "【管吊架】";
                //利用DateTime判斷寫入資料的欄位是否要變更
                excelWorksheet.Cells[1, 1] = theDate;
                excelWorksheet.Cells[1, 2] = Counter.count;

                //excelApp.ActiveWorkbook.SaveAs(@"C:\Users\12061753\Desktop\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                excelApp.ActiveWorkbook.SaveAs(finalPath, Excel.XlSaveAsAccessMode.xlNoChange);
                //excelWorkbook.SaveAs(@"C:\Users\12061753\Documents\abc.xls");

                //關閉及釋放物件
                //excelWorksheet = null;
                //excelWorkbook.Close();
                //excelWorkbook = null;
                //excelApp.Quit();
                //excelApp = null;        

            }
        }
    }
    class pipeHanger
    {
        //1.自動匯入元件檔的功能
        //2.自動判斷大小後選擇欲放置的吊架
        //public Family pipeHangerFamily(Document doc, double pipeDiameter)

        public FamilySymbol getFamilySymbol(Document doc, double pipeDiameter)
        {
            Family tagetFamily = null;
            string targetFamilyName = PIpeHangerSetting.Default.FamilySelected;
            ElementFilter FamilyFilter = new ElementClassFilter(typeof(Family));
            FilteredElementCollector hangerCollector = new FilteredElementCollector(doc);
            hangerCollector.WherePasses(FamilyFilter).ToElements();
            foreach (Family family in hangerCollector)
            {
                if (family.Name == targetFamilyName)
                {
                    tagetFamily = family;
                }
            }
            if (tagetFamily == null) MessageBox.Show("尚未設定要使用的吊架!!");

            //以管徑判斷，取得targetFamily下的管徑
            FamilySymbol targetSymbol = null;
            if (tagetFamily != null)
            {
                foreach (ElementId hangId in tagetFamily.GetFamilySymbolIds())
                {
                    FamilySymbol tempSymbol = doc.GetElement(hangId) as FamilySymbol;
                    if (targetFamilyName == "M_光纖纜架_管附件")
                    {
                        targetSymbol = tempSymbol;
                    }
                    else
                    {
                        double hangerDiameter = tempSymbol.LookupParameter("標稱直徑").AsDouble(); //利用標稱直徑的參數作為判斷依據
                        if (hangerDiameter == pipeDiameter)
                        {
                            targetSymbol = tempSymbol;
                        }
                    }
                }
                if (targetSymbol == null)
                {
                    MessageBox.Show("預設的吊架沒有和管匹配的類型，請重新設定!!");
                }
            }
            return targetSymbol;
        }
        public FamilyInstance CreateHanger(Document doc, XYZ location, Element element, FamilySymbol targetFamily)
        {
            //創造吊架
            FamilyInstance instance = null;
            if (targetFamily != null)
            {
                targetFamily.Activate();
                if (null != targetFamily)
                {
                    MEPCurve pipCrv = element as MEPCurve; //選取管件，一定可以轉型MEPCurve
                    Level hangLevel = pipCrv.ReferenceLevel;
                    Level targetLevel = null;
                    //針對連結管，取得樓層的方式需要微調
                    string levelName = hangLevel.Name;
                    FilteredElementCollector tempCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType();
                    foreach (Element e in tempCollector)
                    {
                        if (e.Name == levelName)
                        {
                            Level temp = e as Level;
                            targetLevel = temp;
                        }
                    }
                    if(targetLevel==null)
                    {
                        MessageBox.Show("請確認本機檔與連結檔中的樓層名稱是否一致");
                    }
                    double moveDown = targetLevel.ProjectElevation; //取得該層樓高層
                    instance = doc.Create.NewFamilyInstance(location, targetFamily, targetLevel, StructuralType.NonStructural); //一定要宣告structural 類型? yes
                    double toMove2 = location.Z - moveDown;
#if RELEASE2019
                    instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(toMove2); //因為給予instance reference level後，實體會基於level的高度上進行偏移，因此需要將偏移量再扣掉一次，非常重要 !!!!。
#else
                    instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(toMove2); //因為給予instance reference level後，實體會基於level的高度上進行偏移，因此需要將偏移量再扣掉一次，非常重要 !!!!。
#endif
                }
            }
            return instance;
        }
    }
    public class PipeSelectionFilter : Autodesk.Revit.UI.Selection.ISelectionFilter
    {
        private Document _doc;
        public PipeSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            if (element.Category.Id == pipe.Id)
            {
                return true;
            }
            else if (element.Category.Id == duct.Id)
            {
                return true;
            }
            else if (element.Category.Id == conduit.Id)
            {
                return true;
            }
            else if (element.Category.Id == tray.Id)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return true;
        }
    }
    public class linkedPipeSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public linkedPipeSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            return true;
        }
        public bool AllowReference(Reference refer, XYZ point)
        {
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            var elem = this._doc.GetElement(refer);
            if (elem != null && elem is RevitLinkInstance link)
            {
                var linkElem = link.GetLinkDocument().GetElement(refer.LinkedElementId);
                if (linkElem.Category.Id == pipe.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == duct.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == conduit.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == tray.Id)
                {
                    return true;
                }
            }
            else
            {
                if (elem.Category.Id == pipe.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == duct.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == conduit.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == tray.Id)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
