using Npgsql;
using EasyReportService.Entities;
using EasyReportService.Utils;
using NLog;
using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;

namespace EasyReportService.Services
{
    public class EasyReportService
    {
        #region 基本設定檔
        private readonly ReportSetting reportSetting;
        private readonly MailSetting mailSetting;
        private readonly Logger logger;
        private readonly string connectionStrings;
        #endregion 基本設定檔

        public EasyReportService(ReportSetting reportSettings, MailSetting mailSettings)
        {
            this.reportSetting = reportSettings;
            this.mailSetting = mailSettings;
            this.logger = NLog.LogManager.GetCurrentClassLogger();
            this.connectionStrings = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ToString();
            reportSetting.logger = this.logger;
        }

        public void Run(ReportData reportData = null)
        {
            try
            {
                writeLog("InitialFolder Start --");
                InitialFolder();
                writeLog("InitialFolder Done --");

                writeLog("GetQuery Start --");
                var reportDataLst = (reportData == null || string.IsNullOrEmpty(reportData.Command)) ?
                    GetQuery().ToList() : new List<ReportData> { reportData };
                writeLog("GetQuery Done --");

                DoExportAndSendMail(reportDataLst);
            }
            catch (Exception ex)
            {
                logger.Error(ex.ToString());
            }
            writeLog("Process Done --");
        }

        /// <summary> 初始化Folder </summary>
        private void InitialFolder()
        {
            //check folder exist
            DirectoryInfo export = new DirectoryInfo(reportSetting.ExportPH);
            if (!export.Exists)
            {
                export.Create();
            }
            DirectoryInfo query = new DirectoryInfo(reportSetting.QueryPH);
            if (!query.Exists)
            {
                query.Create();
            }
        }

        ///<summary> 取得SQL查詢指令 </summary>
        /// <returns></returns>
        private IEnumerable<ReportData> GetQuery()
        {
            var reportLst = new List<ReportData>();
            //todo get directory FileList
            var queryFileLst = System.IO.Directory.GetFiles(reportSetting.QueryPH);

            if (queryFileLst.Count() > 0)
            {
                foreach (var item in queryFileLst)
                {
                    var exportFileName = Path.GetFileNameWithoutExtension(item);
                    var command = System.IO.File.ReadAllText(item);
                    reportLst.Add(new ReportData { FileName = exportFileName, Command = command });
                }
            }
            return reportLst;
        }

        /// <summary> 執行SQL查詢指令 </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        private DataSet RunQuery(string command)
        {
            switch (reportSetting.SqlType)
            {
                case SQLProviderType.MSSQL:
                    return RunMssqlQuery(command);

                case SQLProviderType.POSTGRE:
                    return RunPostgreQuery(command);

                case SQLProviderType.ORACLE:
                    return RunOracleQuery(command);

                default:
                    break;
            }
            return null;
        }

        /// <summary> 執行Postgre Query </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        private DataSet RunPostgreQuery(string command)
        {
            using (var conn = new NpgsqlConnection(connectionStrings))
            {
                using (var cmd = new NpgsqlCommand(command, conn))
                {
                    using (NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd))
                    {
                        DataSet ds = new DataSet();
                        da.Fill(ds);
                        if (ds.Tables.Count > 0)
                        {
                            return ds;
                        }
                        return null;
                    }
                }
            }
        }

        /// <summary> 執行MSSQL Query </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        private DataSet RunMssqlQuery(string command)
        {
            using (var conn = new SqlConnection(connectionStrings))
            {
                using (var cmd = new SqlCommand(command, conn))
                {
                    using (var da = new SqlDataAdapter(cmd))
                    {
                        var ds = new DataSet();
                        da.Fill(ds);
                        if (ds.Tables.Count > 0)
                        {
                            return ds;
                        }
                        return null;
                    }
                }
            }
        }

        /// <summary>
        /// 執行 ORACLE Query
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        private DataSet RunOracleQuery(string command)
        {
            using (var conn = new OracleConnection(connectionStrings))
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = command;
                    using (OracleDataAdapter da = new OracleDataAdapter(cmd))
                    {
                        var ds = new DataSet();
                        da.Fill(ds);
                        if (ds.Tables.Count > 0)
                        {
                            return ds;
                        }
                        return null;
                    }
                }
            }
        }

        /// <summary> 執行匯出與寄信 </summary>
        /// <param name="reportDataLst"></param>
        private void DoExportAndSendMail(IEnumerable<ReportData> reportDataLst)
        {
            foreach (var item in reportDataLst)
            {
                writeLog(string.Format("RunQuery [{0}] Start --", item.FileName));
                var data = RunQuery(item.Command);
                var fileName = string.Empty;
                writeLog("RunQuery Done --");

                if (data != null && data.Tables[0].Columns.Count > 0)
                {
                    writeLog(string.Format("ExportExcel [{0}] Start --", item.FileName));
                    fileName = ExportData(data, item.FileName);
                    writeLog("ExportExcel Done --");

                    if (reportSetting.NeedSend)
                    {
                        writeLog(string.Format("SendMail [{0}] Start --", item.FileName));
                        SendMailing(fileName);
                        writeLog("SendMail Done --");
                    }
                }
                else
                {
                    writeLog(string.Format("{0}, Command:{1}", "No Result...", item.Command));
                }
            }
        }

        /// <summary> 匯出報表 - EXCEL </summary>
        /// <param name="data"></param>
        /// <param name="exportFileName"></param>
        /// <returns></returns>
        private string ExportData(DataSet data, string exportFileName)
        {
            var fileName = string.Empty;
            switch (reportSetting.ExportType)
            {
                default:
                case ExportType.CSV:
                    fileName = string.Format("{0}_{1}.csv", exportFileName, DateTime.Now.ToString("yyyyMMddHHmmss"));
                    CsvUtils.export(data, reportSetting.ExportPH, fileName);
                    break;

                case ExportType.EXCEL:
                    fileName = string.Format("{0}_{1}.xls", exportFileName, DateTime.Now.ToString("yyyyMMddHHmmss"));
                    ExcelUtils.export(data, reportSetting.ExportPH, fileName);
                    break;
            }
            return fileName;
        }

        /// <summary> 夾帶寄信 </summary>
        /// <param name="fileName"></param>
        private void SendMailing(string fileName)
        {
            try
            {
                var sendmail = new SendMail(reportSetting, mailSetting);
                var subject = fileName;
                var msg = "請參閱附件，謝謝";

                Attachment file =
                    AttachmentHelper.CreateAttachment(Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, reportSetting.ExportPH, fileName), fileName, TransferEncoding.Base64);
                sendmail.send_email(msg, subject, mailSetting.MailTo, mailSetting.Mailcc, mailSetting.MailBcc, file);
            }
            catch (Exception ex)
            {
                writeLog(ex.ToString());
            }
        }

        /// <summary> 記錄Log </summary>
        /// <param name="msg"></param>
        private void writeLog(string msg)
        {
            System.Console.WriteLine(msg);
            logger.Info(msg);
        }

    }
}
