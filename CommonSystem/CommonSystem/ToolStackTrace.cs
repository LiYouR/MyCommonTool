using System;

namespace CommonSystem
{
    public static class ToolStackTrace
    {
        public static string GetStackTrace()
        {
            string strRet = string.Empty;
            System.Diagnostics.StackTrace st = new System.Diagnostics.StackTrace();
            System.Diagnostics.StackFrame[] sfs = st.GetFrames();
            if (null != sfs)
            {
                for (int i = 1; i < sfs.Length; ++i)
                {
                    strRet += sfs[i].GetMethod().ToString();
                    strRet += "  Line:" + sfs[i].GetFileLineNumber();
                    strRet += "\r\n";
                }
            }
            return strRet;
        }

        public const string LOGFORMAT_EXCEPTION = "Exception({3}) Message：{0}，\r\nSource：{1}，\r\nStackTrace：{2}，\r\nMethod：{4}\r\n";

        public static string GetStackTrace(Exception e)
        {
            string strRet = string.Format(LOGFORMAT_EXCEPTION, e.Message, e.Source, e.StackTrace, e.GetType(), e.TargetSite.Name);
            return strRet;
        }
    }
}
