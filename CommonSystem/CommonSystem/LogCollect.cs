using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.IO.Compression;
using System.Threading;
using CommonTool;
using CommonTool.IOEx;

namespace CommonSystem
{
    public class LogCollect
    {
        //Target Type
        public enum LogKind
        {
            Client,
            Server
        }

        //CultureInfo
        private static CultureInfo _cl = CultureInfo.CreateSpecificCulture("en-US");

        //PATH
        private static readonly string APP_DATA_PATH = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        private static readonly string TOPSS_CLIENT_LOG_PATH = Path.Combine(APP_DATA_PATH, @"Sony\Topss\Log");
        private static readonly string TOPSS_SERVER_LOG_PATH = Path.Combine(APP_DATA_PATH, @"Sony\TOPSSServer\Log");
        private static readonly string TOPSS_SERVER_DB_LOG_PATH = Path.Combine(APP_DATA_PATH, @"Sony\TOPSSServer\DBLog");
        private static readonly string TOPSS_CLIENT_XML_PATH = Path.Combine(APP_DATA_PATH, @"Sony\Topss");
        private static readonly string TOPSS_SERVER_XML_PATH = Path.Combine(APP_DATA_PATH, @"Sony\TOPSSServer");
        private static readonly string TOPSS_CLIENT_TEMP_PATH = Path.Combine(APP_DATA_PATH, @"Sony\Topss\Temp");
        private static readonly string TOPSS_SERVER_TEMP_PATH = Path.Combine(APP_DATA_PATH, @"Sony\TOPSSServer\Temp");

        private static readonly string EXPORT_ROOT_SYSTEM_LOG = @"SystemLog";
        private static readonly string EXPORT_ROOT_LOG = @"Log";
        private static readonly string EXPORT_ROOT_DB_LOG = @"DBLog";
        private static readonly string EXPORT_ROOT_XML = @"Xml";
        private static readonly string EXPORT_ROOT_DUMP = @"Dump";

        //Member
        private DateTime _dateFrom = DateTime.Now;
        private DateTime _dateTo = DateTime.Now;
        private string _strSrcPath = "";
        private string _strDestPath = "";

        //Client Export Target
        private readonly SystemEventLog.SystemEventKind[] TopssClientSystemLogTarget =
        {
            SystemEventLog.SystemEventKind.Application,
            SystemEventLog.SystemEventKind.System,
            SystemEventLog.SystemEventKind.Security,
            SystemEventLog.SystemEventKind.Setup
        };

        //Server Export Target
        private readonly SystemEventLog.SystemEventKind[] TopssServerSystemLogTarget =
        {
            SystemEventLog.SystemEventKind.Application,
            SystemEventLog.SystemEventKind.System,
            SystemEventLog.SystemEventKind.Security,
            SystemEventLog.SystemEventKind.Setup,
            SystemEventLog.SystemEventKind.TOPSSServices
        };

        /// <summary>
        /// Log Collect
        /// </summary>
        /// <param name="logKind"></param>
        /// <param name="strExportPath"></param>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <param name="actionProgress"></param>
        /// <returns>0:OK -1:Failed -2:Cancel</returns>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public int Collect(LogKind logKind, string strExportPath, DateTime dateFrom, DateTime dateTo, Action<int> actionProgress, CancellationTokenSource cancelToken)
        {
            if (actionProgress != null) actionProgress(5);

            //Init
            _dateFrom = dateFrom;
            _dateTo = dateTo;

            string strTempPath = "";
            SystemEventLog.SystemEventKind[] arrayExportTarget = null;
            string strLogPath = "";
            string strDBLogPath = "";
            string strXmlPath = "";

            if (logKind == LogKind.Client)
            {
                strTempPath = TOPSS_CLIENT_TEMP_PATH;
                arrayExportTarget = TopssClientSystemLogTarget;
                strLogPath = TOPSS_CLIENT_LOG_PATH;
                strDBLogPath = "";
                strXmlPath = TOPSS_CLIENT_XML_PATH;
            }
            else
            {
                strTempPath = TOPSS_SERVER_TEMP_PATH;
                arrayExportTarget = TopssServerSystemLogTarget;
                strLogPath = TOPSS_SERVER_LOG_PATH;
                strDBLogPath = TOPSS_SERVER_DB_LOG_PATH;
                strXmlPath = TOPSS_SERVER_XML_PATH;
            }

            string TempFolderName = MakeTempFolderName(logKind,dateFrom, dateTo);
            strTempPath = Path.Combine(strTempPath, TempFolderName);

            if (actionProgress != null) actionProgress(10);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }


            //Clear Temp Folder
            if (Directory.Exists(strTempPath))
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
            }
            Directory.CreateDirectory(strTempPath);

            if (actionProgress != null) actionProgress(15);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //System Log Collect
            string strSystemLogPath = Path.Combine(strTempPath, EXPORT_ROOT_SYSTEM_LOG);
            foreach (var eventKind in arrayExportTarget)
            {
                SystemEventLog.ExportSystemEvent(eventKind, strSystemLogPath, dateFrom, dateTo);
            }

            if (actionProgress != null) actionProgress(40);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //Log Collect
            if (!string.IsNullOrEmpty(strLogPath))
            {
                _strSrcPath = strLogPath;
                _strDestPath = Path.Combine(strTempPath, EXPORT_ROOT_LOG);
                CommonTool.DirectoryEx.TraversalDirectory(strLogPath, DoLogFileCollect);
            }

            if (actionProgress != null) actionProgress(55);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //DB Log Collect
            if (!string.IsNullOrEmpty(strDBLogPath))
            {
                _strSrcPath = strDBLogPath;
                _strDestPath = Path.Combine(strTempPath, EXPORT_ROOT_DB_LOG);
                CommonTool.DirectoryEx.TraversalDirectory(strDBLogPath, DoLogFileCollect);
            }

            if (actionProgress != null) actionProgress(70);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //XML Collect
            if (!string.IsNullOrEmpty(strXmlPath))
            {
                _strSrcPath = strXmlPath;
                _strDestPath = Path.Combine(strTempPath, EXPORT_ROOT_XML);
                CommonTool.DirectoryEx.TraversalDirectory(_strSrcPath, DoXMLFileCollect, false);
            }

            if (actionProgress != null) actionProgress(80);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //Dump Collect
            if (!string.IsNullOrEmpty(strLogPath))
            {
                _strSrcPath = strLogPath;
                _strDestPath = Path.Combine(strTempPath, EXPORT_ROOT_DUMP);
                CommonTool.DirectoryEx.TraversalDirectory(_strSrcPath, DoDumpFileCollect, false);
            }
            if (actionProgress != null) actionProgress(90);
            if (cancelToken.IsCancellationRequested)
            {
                CommonTool.DirectoryEx.DeleteDirectory(strTempPath);
                return -2;
            }

            //Compress To ZIP
            if (!strExportPath.EndsWith(".zip")) strExportPath += ".zip";
            //string zipPath = strExportPath + @"\" + TempFolderName + ".zip";
            ZIP.Compress(strTempPath, strExportPath);

            //Delete Temp File 
            CommonTool.DirectoryEx.DeleteDirectory(strTempPath);

            if (cancelToken.IsCancellationRequested)
            {
                return -2;
            }
            if (actionProgress != null) actionProgress(100);


            return 0;
        }

        private string MakeTempFolderName(LogKind logKind,DateTime dateFrom, DateTime dateTo)
        {
            return Enum.GetName(typeof(LogKind), logKind) + "_" +dateFrom.ToString("yyyyMMdd") + "_" + dateTo.ToString("yyyyMMdd");
        }

        private void DoLogFileCollect(string strSrcFilePath)
        {
            if (!File.Exists(strSrcFilePath))
            {
                return;
            }

            if (!Directory.Exists(_strDestPath))
            {
                Directory.CreateDirectory(_strDestPath);
            }

            FileInfo fileInfo = new FileInfo(strSrcFilePath);

            if (fileInfo.Extension == ".log")
            {
                DateTime fileDate = new DateTime();
                string strBuf = System.Text.RegularExpressions.Regex.Match(fileInfo.Name, @"((19|20)\d{2})([0][1-9]|[1][0-2])([0-2][0-9]|[3][0-1])").ToString();
                if (string.IsNullOrEmpty(strBuf) || !DateTime.TryParseExact(strBuf, "yyyyMMdd", _cl, DateTimeStyles.None, out fileDate))
                {
                    fileDate = fileInfo.CreationTime;
                }

                if (fileDate >= _dateFrom && fileDate <= _dateTo)
                {
                    if (fileInfo.DirectoryName != null)
                    {
                        string strDestDir = fileInfo.DirectoryName.Replace(_strSrcPath, _strDestPath);


                        if (!Directory.Exists(strDestDir))
                        {
                            Directory.CreateDirectory(strDestDir);
                        }

                        File.Copy(fileInfo.FullName, Path.Combine(strDestDir, fileInfo.Name));
                    }
                }
            }
        }

        private void DoXMLFileCollect(string strSrcFilePath)
        {
            if (!File.Exists(strSrcFilePath))
            {
                return;
            }

            if (!Directory.Exists(_strDestPath))
            {
                Directory.CreateDirectory(_strDestPath);
            }

            FileInfo fileInfo = new FileInfo(strSrcFilePath);

            if (fileInfo.Extension == ".xml")
            {
                if (fileInfo.DirectoryName != null)
                {
                    string strDestDir = fileInfo.DirectoryName.Replace(_strSrcPath, _strDestPath);


                    if (!Directory.Exists(strDestDir))
                    {
                        Directory.CreateDirectory(strDestDir);
                    }

                    File.Copy(fileInfo.FullName, Path.Combine(strDestDir, fileInfo.Name));
                }
            }
        }

        private void DoDumpFileCollect(string strSrcFilePath)
        {
            if (!File.Exists(strSrcFilePath))
            {
                return;
            }

            if (!Directory.Exists(_strDestPath))
            {
                Directory.CreateDirectory(_strDestPath);
            }

            FileInfo fileInfo = new FileInfo(strSrcFilePath);

            if (fileInfo.Extension == ".dmp")
            {
                if (fileInfo.DirectoryName != null)
                {
                    string strDestDir = fileInfo.DirectoryName.Replace(_strSrcPath, _strDestPath);


                    if (!Directory.Exists(strDestDir))
                    {
                        Directory.CreateDirectory(strDestDir);
                    }

                    File.Copy(fileInfo.FullName, Path.Combine(strDestDir, fileInfo.Name));
                }
            }
        }
    }
}
