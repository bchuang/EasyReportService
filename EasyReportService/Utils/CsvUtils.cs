using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyReportService.Utils
{
    public class CsvUtils
    {
         /// <summary> 匯出CSV報表 </summary>
        /// <param name="exceldata"> 匯出資料集 </param>
        /// <param name="savePath"> 儲存路徑 </param>
        /// <param name="fileName"> 儲存檔名 </param>
        public static void export(DataSet exceldata, string savePath, string fileName)
        {
            string data = "";
            var oTable = exceldata.Tables[0];
            StreamWriter wr = new StreamWriter(Path.Combine(savePath, fileName), false, System.Text.Encoding.Default);

            foreach (DataColumn column in oTable.Columns)
            {
                data += column.ColumnName + ",";
            }
            data += "\n";
            wr.Write(data);
            data = "";

            foreach (DataRow row in oTable.Rows)
            {
                foreach (DataColumn column in oTable.Columns)
                {
                    data += row[column].ToString().Trim() + ",";
                }
                data += "\n";
                wr.Write(data);
                data = "";
            }
            data += "\n";

            wr.Dispose();
            wr.Close();
        }
    }
}
