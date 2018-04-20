using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace CommonSystem
{
    public class SystemEventLog
    {
        public enum SystemEventKind
        {
            Application = 0,
            Security,
            Setup,
            System,
            TOPSSServices
        }

        private static readonly string[] _systemEventKindName =
        {
            "Application",
            "Security",
            "Setup",
            "System",
            "TOPSS Services"
        };

        private static readonly string[] _ExportFileName =
        {
            "Application.evtx",
            "Security.evtx",
            "Setup.evtx",
            "System.evtx",
            "TOPSS_Services.evtx"
        };

        private const string _DATETIME_FORMAT = "yyyy-MM-ddTHH:mm:ss.fffffff00Z";

        private const string _EVENT_SEARCH_CONDITION = "*[System/TimeCreated[@SystemTime>\"{0}\" and @SystemTime<\"{1}\"]]";

        public static bool ExportSystemEvent(SystemEventKind eventKind, string strExportPath, DateTime dateFrom, DateTime dateTo)
        {
            string strDateFrom = dateFrom.ToUniversalTime().ToString(_DATETIME_FORMAT);
            string strDateTo = dateTo.ToUniversalTime().ToString(_DATETIME_FORMAT);
            string strExportFileName = _ExportFileName[(int)eventKind];
            string strFullPath = Path.Combine(strExportPath, strExportFileName);

            try
            {
                if (!Directory.Exists(strExportPath))
                {
                    Directory.CreateDirectory(strExportPath);
                }

                if (File.Exists(strFullPath))
                {
                    File.Delete(strFullPath);
                }

                string queryStr = string.Format(_EVENT_SEARCH_CONDITION, strDateFrom, strDateTo);

                EventLogSession els = new EventLogSession();
                els.ExportLogAndMessages(
                                    _systemEventKindName[(int)eventKind],        //  Log Name to archive
                                    PathType.LogName,                            //  Type of Log
                                    queryStr,                                    //  Query selecting events
                                    strFullPath,                                 //  Exported Log Path(log created by this operation) 
                                    false,                                       //  Stop the archive if the query is invalid
                                    CultureInfo.CurrentCulture);                 //  Culture to archive the events in

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }

            return true;
        }
    }
}
