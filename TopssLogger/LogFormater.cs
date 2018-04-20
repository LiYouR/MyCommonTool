using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace TopssLogger
{
    public static class LogFormater
    {
        private const string LOGFORMAT_EXCEPTION = "************** Exception Text **************\r\n{0} : {1}\r\n{2}";
        private const string LOGFORMAT_BASE = "[{0,5}] [{1}] [{2,2}] {3} : {4}";
        private const string LOGFORMAT_STACK = "[{0}] [{1}] [Line:{2,5}]";
        private const string LOGFORMAT_DATETIME = "yyyy/MM/dd HH:mm:ss.fff";

        /// <summary>
        /// FormatLog
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="strMessage"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static string FormatLog(LogLevel logType, string strMessage, Exception e)
        {
            //異常情報
            if (e!= null)
            {
                string strException = string.Format(LOGFORMAT_EXCEPTION, e.GetType(), e.Message, e.StackTrace);
                strMessage = strMessage + "\r\n" + strException;
            }

            //StackTrace情報
            StackTrace st = new StackTrace(true);
            StackFrame sf = st.GetFrame(2);
            string strStackinfo = string.Format(LOGFORMAT_STACK, Path.GetFileName(sf.GetFileName()), sf.GetMethod().Name,
                sf.GetFileLineNumber());

            //Format
            string s = string.Format(LOGFORMAT_BASE, logType.ToString().ToUpper(),
                    DateTime.Now.ToString(LOGFORMAT_DATETIME), Thread.CurrentThread.ManagedThreadId, strStackinfo, strMessage);

            return s;
        }
    }
}
