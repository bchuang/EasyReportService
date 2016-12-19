using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace EasyReportService.Utils
{
    public class ExcelUtils
    {
        /// <summary> 匯出Excel報表 </summary>
        /// <param name="exceldata"> 匯出資料集 </param>
        /// <param name="savePath"> 儲存路徑 </param>
        /// <param name="fileName"> 儲存檔名 </param>
        /// <param name="titleLst"> 欄名清單 </param>
        public static void export(DataSet exceldata, string savePath, string fileName)
        {
            Application App = null;
            Workbook Wbook = null;
            Worksheet Wsheet = null;

            try
            {
                App = new Application();
                Wbook = App.Workbooks.Add();
                Wsheet = (Worksheet)Wbook.ActiveSheet;
                int colNum = 1;
                //報表抬頭
                foreach (DataColumn item in exceldata.Tables[0].Columns)
                {
                    Wsheet.Cells[1, colNum].value = item.ColumnName;
                    colNum++;
                }

                //報表內容
                int colCount = exceldata.Tables[0].Columns.Count;
                var i = 2;
                if (exceldata != null)
                {
                    var data = exceldata.Tables[0];
                    for (int n = 0; n < data.Rows.Count; n++)
                    {
                        for (int col = 0; col < colCount; col++)
                        {
                            Wsheet.Cells[i, col + 1].NumberFormat = "@";
                            Wsheet.Cells[i, col + 1].value = data.Rows[n][col];
                        }
                        i++;
                    }
                }

                for (int col = 0; col < colCount; col++)
                {
                    Wsheet.Columns[col + 1].AutoFit();
                }

                //設置禁止彈出保存和覆蓋的詢問提示框
                Wsheet.Application.DisplayAlerts = false;
                Wsheet.Application.AlertBeforeOverwriting = false;

                //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
                //Wbook.Save();

                //另存活頁簿
                //string savePath = ConfigurationManager.AppSettings["exportPath"];
                Wbook.SaveAs(Path.Combine(savePath, fileName), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:" + ex.Message);
            }
            finally
            {
                if (Wsheet != null)
                {
                    Marshal.FinalReleaseComObject(Wsheet);
                }
                if (Wbook != null)
                {
                    Wbook.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
                    Marshal.FinalReleaseComObject(Wbook);
                }
                if (App != null)
                {
                    App.Workbooks.Close();
                    App.Quit();
                    Marshal.FinalReleaseComObject(App);
                }

                //關閉EXCEL
                //Wbook.Close();

                //離開應用程式
                //App.Quit();
            }
        }
    }
}