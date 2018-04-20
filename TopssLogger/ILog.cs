using System;

namespace TopssLogger
{
    public interface ILog
    {
        /// <summary>
        /// ログを書き込む
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="strMessage"></param>
        /// <param name="e"></param>
        void Write(LogLevel logType, string strMessage, Exception e);

        /// <summary>
        /// ログを書き込む
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="strMessage"></param>
        void Write(LogLevel logType, string strMessage);

        /// <summary>
        /// メソッド開始
        /// </summary>
        /// <param name="strMessage"></param>
        void WriteMethodStart(string strMessage);

        /// <summary>
        /// メソッド終了
        /// </summary>
        /// <param name="strMessage"></param>
        void WriteMethodEnd(string strMessage);
    }
}
