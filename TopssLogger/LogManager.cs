using System;
using System.Collections.Concurrent;

namespace TopssLogger
{
    public sealed class LogManager
    {
        //出力ルートパス
        private static string _logRootPath = "";

        //ディフォルト出力インスタンス　キー
        private const string DefaultLoggerKey = "Default";

        //出力インスタンス
        private static ConcurrentDictionary<string, ILog> _dicLoggers = new ConcurrentDictionary<string, ILog>();
        
        //Locker
        private static object _lock = new object();

        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="strLogRootPath"></param>
        public static void InitLogManager(string strLogRootPath)
        {
            _logRootPath = strLogRootPath;
        }

        /// <summary>
        /// インスタンス取得
        /// </summary>
        /// <returns></returns>
        public static ILog GetLogger()
        {
            if (!_dicLoggers.ContainsKey(DefaultLoggerKey))
            {
                lock (_lock)
                {
                    if (!_dicLoggers.ContainsKey(DefaultLoggerKey))
                    {
                        TopssLogger topssLoger = new TopssLogger(_logRootPath, "");
                        topssLoger.SetLogOutputLevel(LogLevel.All, true);
                        _dicLoggers[DefaultLoggerKey] = topssLoger;
                    }
                }
            }
            return _dicLoggers[DefaultLoggerKey];
        }

        /// <summary>
        /// 名前でインスタンス取得
        /// </summary>
        /// <param name="strLogInstanceName"></param>
        /// <returns></returns>
        public static ILog GetLogger(String strLogInstanceName)
        {
            if (!_dicLoggers.ContainsKey(strLogInstanceName))
            {
                lock (_lock)
                {
                    if (!_dicLoggers.ContainsKey(strLogInstanceName))
                    {
                        TopssLogger topssLoger = new TopssLogger(_logRootPath, strLogInstanceName);
                        topssLoger.SetLogOutputLevel(LogLevel.All, true);
                        _dicLoggers[strLogInstanceName] = topssLoger;
                    }
                }
            }
            return _dicLoggers[strLogInstanceName];
        }

        /// <summary>
        /// 出力ログラベル設定
        /// </summary>
        /// <param name="outputLogLevel"></param>
        /// <param name="bOutput"></param>
        public static void SetLogLevel(LogLevel outputLogLevel, bool bOutput)
        {
            TopssLogger topssLoger = (TopssLogger)GetLogger();
            topssLoger.SetLogOutputLevel(outputLogLevel, bOutput);
        }

        /// <summary>
        /// 名前で指定インスタンスの出力ログラベル設定
        /// </summary>
        /// <param name="strLogInstanceName"></param>
        /// <param name="outputLogLevel"></param>
        /// <param name="bOutput"></param>
        public static void SetLogLevel(String strLogInstanceName, LogLevel outputLogLevel, bool bOutput)
        {
            TopssLogger topssLoger = (TopssLogger)GetLogger(strLogInstanceName);
            topssLoger.SetLogOutputLevel(outputLogLevel, bOutput);
        }

        /// <summary>
        /// ShutDown
        /// </summary>
        public static void ShutDown()
        {
            lock (_lock)
            {
                foreach (ILog loger in _dicLoggers.Values)
                {
                    ((TopssLogger)loger).Dispose();
                }
                _dicLoggers.Clear();
            }
        }
    }
}
