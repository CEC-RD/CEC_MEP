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


namespace BeamCasing_ButtonCreate
{
    public static class Counter
    {
        public static int count = 0;
    }
    public class ExcelLog
    {
        //private string _filePath;
        //private string _fileName;
        private int _userCount;

        public enum ribbons
        {
            autoHanger = 1,
            beamCast = 2,
            wallCast = 3,
            blockTrans = 4,
            perFab = 5
        }
        public ExcelLog(int count)
        {
            //this._filePath = filePath;
            this._userCount = count;
        }
        //建立Enum匹配Ribbon Name
        public void userLog()
        {
            //根據不同支外掛寫入不同excel
            //string name = ribbons.beamCast.ToString();
            int ad = 1;//格式校正
            string userName = Environment.UserName;
            string fileName = userName + ".xls";
            string ribbonName = ribbons.beamCast.ToString();

            int startYear = 2021;//開始紀錄的前一年
            int year = DateTime.Today.Year;
            int month = DateTime.Today.Month + ad;

            //寫入Excel
            DateTime toDay = DateTime.Today.Date;
            Excel.Application excelApp = new Excel.Application();
            //找到dropox的位置
            var infoPath = @"Dropbox\info.json";
            var jsonPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), infoPath);
            if (!File.Exists(jsonPath)) jsonPath = Path.Combine(Environment.GetEnvironmentVariable("AppData"), infoPath);
            if (!File.Exists(jsonPath)) throw new Exception("請安裝並登入Dropbox桌面應用程式!");
            var dropboxPath = File.ReadAllText(jsonPath).Split('\"')[5];
            //var filePath = @"\BIM\05 Common 共通\Source\Revit外掛\CEC MEP\00.Journal logs\";
            string filepath = @"N:\CEC_API\logs\";
            string finalPath = filepath + fileName;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelWorksheet = null;
            //利用DateTime判斷寫入資料的欄位是否要變更
            int columNum = year - startYear + ad;
            bool fileExist = false;
            bool sheetExist = false;
            try
            {
                if (excelApp != null)
                {
                    if (File.Exists(finalPath))
                    {

                        fileExist = true;
                        //打開既有的檔案，判斷工作簿是否存在
                        excelWorkbook = excelApp.Workbooks.Open(finalPath);
                        foreach (Excel.Worksheet sheet in excelWorkbook.Worksheets)
                        {
                            if (sheet.Name == ribbonName)
                            {
                                sheetExist = true;
                                excelWorksheet = sheet;
                            }
                        }
                        //if (sheet.Name == year.ToString())
                        //如果Excelsheet已存在
                        if (sheetExist == true)
                        {
                            //把資料寫進相對的workbook
                            excelWorksheet.Cells[1, columNum] = year.ToString() + "年";
                            //必須針對這個判斷時否有值
                            var cell = excelWorksheet.Cells[month, columNum] as Excel.Range;
                            var cellValue = 0;
                            if (cell.Value == null)
                            {
                                cellValue = 0;
                            }
                            else
                            {
                                cellValue = (int)(excelWorksheet.Cells[month, columNum] as Excel.Range).Value;
                            }
                            excelWorksheet.Cells[month, columNum] = cellValue + _userCount;
                        }
                        else
                        {
                            //excelWorksheet = new Excel.Worksheet();
                            //excelWorkbook.Worksheets.Add();
                            //int worksheetsNum = excelWorkbook.Worksheets.Count;
                            //excelWorksheet = excelWorkbook.Worksheets[1];
                            excelWorksheet = excelWorkbook.Worksheets.Add();
                            excelWorksheet.Name = ribbonName;
                            //excelWorksheet.Name = year.ToString();
                            //建立一個新的以年份為基礎的workbook
                            //把資料寫進相對的sheet
                            for (int i = 1; i <= 12; i++)
                            {
                                excelWorksheet.Cells[i + 1, 1] = $"{i}月";
                            }
                            excelWorksheet.Cells[1, columNum] = year.ToString() + "年";
                            excelWorksheet.Cells[month, columNum] = _userCount;
                        }
                    }
                    else
                    {
                        //如果檔案不存在則自己做新的
                        fileExist = false;
                        excelWorkbook = excelApp.Workbooks.Add();
                        excelWorksheet = new Excel.Worksheet();
                        excelWorksheet = excelWorkbook.Worksheets[1];
                        excelWorksheet.Name = ribbonName;
                        for (int i = 1; i <= 12; i++)
                        {
                            excelWorksheet.Cells[i + 1, 1] = $"{i}月";
                        }
                        excelWorksheet.Cells[1, columNum] = year.ToString() + "年";
                        excelWorksheet.Cells[month, columNum] = _userCount;
                    }
                    if (!fileExist)
                    {
                        excelApp.ActiveWorkbook.SaveAs(finalPath, Excel.XlSaveAsAccessMode.xlNoChange);
                    }
                    else
                    {
                        excelApp.ActiveWorkbook.Save();
                    }
                }
                excelWorksheet = null;
                excelWorkbook.Close();
                excelWorkbook = null;
                excelApp.Quit();
                excelApp = null;
            }
            catch
            {
                //關閉及釋放物件
                excelWorksheet = null;
                excelWorkbook.Close();
                excelWorkbook = null;
                excelApp.Quit();
                excelApp = null;
                MessageBox.Show($"{ribbons.wallCast}_log 儲存失敗");
            }
        }
    }
}
