using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyReportService.Entities
{
    public class ReportSetting
    {
        public bool NeedSend { get; set; }
        public bool NeedCmdPause { get; set; }
        public string ExportPH { get; set; }
        public string QueryPH { get; set; }
        public SQLProviderType SqlType { get; set; }
        public ExportType ExportType { get; set; }
        public Logger logger { get; set; }
    }
}
