using EasyReportService.Entities;
using System;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace EasyReportService.Utils
{
    public class SendMail
    {
        #region 基本設定檔
        private readonly ReportSetting reportSetting;
        private readonly MailSetting mailSetting;
        #endregion 基本設定檔

        public SendMail(ReportSetting reportSettings, MailSetting mailSettings)
        {
            this.mailSetting = mailSettings;
            this.reportSetting = reportSettings;
        }

        /// <summary>
        /// 傳入mail主旨與mail內容訊息、收信人
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="mysubject"></param>
        /// <param name="mailto"></param>
        public void send_email(string msg, string mysubject, string mailto)
        {
            try
            {
                var smtpServer = InitialSmtp();
                var message = SetMailMessage(msg, mysubject, mailto);
                smtpServer.Send(message);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(msg);
                reportSetting.logger.Error(ex.ToString());
            }
        }

        /// <summary>
        /// 傳入mail主旨與mail內容訊息、收信人、夾帶檔案
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="mysubject"></param>
        /// <param name="mailto"></param>
        /// <param name="mailcc"></param>
        /// <param name="mailbcc"></param>
        /// <param name="file"></param>
        public void send_email(string msg, string mysubject, string mailto, string mailcc, string mailbcc, Attachment file)
        {
            try
            {
                var smtpServer = InitialSmtp();
                var message = SetMailMessage(msg, mysubject, mailto, mailcc, mailbcc, file);
                smtpServer.Send(message);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(msg);
                reportSetting.logger.Error(ex.ToString());
            }
        }

        /// <summary>
        /// 初始化SMTP
        /// </summary>
        /// <returns></returns>
        private SmtpClient InitialSmtp()
        {
            var smtpServer = new SmtpClient();
            smtpServer.Host = mailSetting.MailServer;
            smtpServer.EnableSsl = mailSetting.EnableSsl;
            smtpServer.Port = mailSetting.Port;
            if (mailSetting.EnableCredential)
            {
                smtpServer.Credentials = new NetworkCredential(mailSetting.MailAccount, mailSetting.MailPwd);
            }
            return smtpServer;
        }

        /// <summary>
        /// 設定信件內容
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="mysubject"></param>
        /// <param name="mailto"></param>
        /// <param name="mailcc"></param>
        /// <param name="mailbcc"></param>
        /// <param name="file"></param>
        /// <returns></returns>
        private MailMessage SetMailMessage(string msg, string mysubject, string mailto,
            string mailcc = null, string mailbcc = null, Attachment file = null)
        {
            var message = new MailMessage();
            message.From = new MailAddress(mailSetting.MailFrom, mysubject, System.Text.Encoding.UTF8);
            message.Subject = mysubject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            message.Body = msg;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = false;
            message.Priority = MailPriority.High;

            ////判斷是否為多個收件者
            var mailtoList = mailto.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            mailtoList.ForEach(mail => message.To.Add(mail));
            ////判斷是否為多個副本者
            var mailccList = mailcc.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            mailccList.ForEach(mail => message.CC.Add(mail));
            ////判斷是否為多個密件副本者
            var mailbccList = mailbcc.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            mailbccList.ForEach(mail => message.Bcc.Add(mail));

            if (file != null)
            {
                message.Attachments.Add(file);
            }
            return message;
        }
    }
}
