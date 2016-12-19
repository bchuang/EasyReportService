using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyReportService.Entities
{
    public class MailSetting
    {
        public string MailServer { get; set; }
        public string MailFrom { get; set; }
        public int Port { get; set; }
        public bool EnableSsl { get; set; }
        public bool EnableCredential { get; set; }
        public string MailAccount { get; set; }
        public string MailPwd { get; set; }
        public string MailTo { get; set; }
        public string Mailcc { get; set; }
        public string MailBcc { get; set; }
    }
}
