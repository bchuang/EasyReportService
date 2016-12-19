using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using EasyReportService.Entities;
using System.Configuration;
using System.IO;

namespace EasyReportService.Services.Tests
{
    [TestClass()]
    public class EasyReportServiceTests
    {
        [TestMethod()]
        public void RunTest()
        {
            GetEasyReportService().Run();
        }

        [TestMethod()]
        public void RunWithReportDataTest()
        {
            GetEasyReportService().Run(new ReportData { FileName = "Test", Command = "SELECT @@SERVERNAME; " });
        }

        private EasyReportService GetEasyReportService()
        {
            return new EasyReportService(GetReportSetting(), GetMailSetting());
        }

        private ReportSetting GetReportSetting()
        {
            return new ReportSetting
            {
                NeedSend = ConfigurationManager.AppSettings["NeedSend"].ToUpper() == "Y" ? true : false,
                NeedCmdPause = ConfigurationManager.AppSettings["NeedCmdPause"].ToUpper() == "Y" ? true : false,
                SqlType = (SQLProviderType)(int.Parse(ConfigurationManager.AppSettings["SQLProviderType"])),
                ExportType = (ExportType)(int.Parse(ConfigurationManager.AppSettings["ExportType"])),
                ExportPH = Path.Combine(string.Format("{0}{1}", System.AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["exportPath"])),
                QueryPH = Path.Combine(string.Format("{0}{1}", System.AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["queryPath"])),
            };
        }

        private MailSetting GetMailSetting()
        {
            return new MailSetting
            {
                MailServer = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailServer"],
                Port = Convert.ToInt32(((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["Port"]),
                EnableSsl = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["EnableSsl"].ToUpper() == "Y" ? true : false,
                MailFrom = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailFrom"],
                MailTo = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailTo"],
                Mailcc = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["Mailcc"],
                MailBcc = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailBcc"],
                EnableCredential = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["EnableCredential"].ToUpper() == "Y" ? true : false,
                MailAccount = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailAccount"],
                MailPwd = ((NameValueCollection)ConfigurationManager.GetSection("mailSettings"))["MailPwd"],
            };
        }

    }
}