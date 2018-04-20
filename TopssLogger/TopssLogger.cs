using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace TopssLogger
{
    public class TopssLogger : ILog, IDisposable
    {
        #region 内部変数

        //ログファイル拡張子名
        private const string LogFileExtension = ".log";

        //ログファイル最大サイズ
        private const uint LogFileSizeLimit = 4;

        //毎日最大ログファイル数
        private const uint LogFileNumPerDayLimit = 8;

        //ログファイルStream
        private StreamWriter _Files;

        //ルートパス 例：D:\Log
        private string _fileRootPath = string.Empty;

        //出力ファイルパス 例：D:\Log\20170814
        private string _filePath = string.Empty;

        //出力ファイル名 例：20170809124500.log
        private string _fileName = string.Empty;

        //出力ファイルフール名 例：D:\Log\20170814\20170809124500.log
        private string _fileFullPath = string.Empty;

        //出力ファイル名の前に、サブ名を付ける。 例：TopssTest_20170814\20170809124500.log
        private string _fileTitle = string.Empty;

        //出力ログのラベル 
        private LogLevel _logOutputLevels = LogLevel.All;

        //Locker
        private object _lockObject = new object();

        //Dispose Flag
        bool _disposed;

        //2017/11/16 Add by DaiJun SPCで、判定画像検索画面で基板シリアルで検索をする際、入力をしても反応が無い場合があり、画像一覧も表示されない場合がある（No.178）[#7788 ] START
        private string pathlog = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                                        @"Sony\TOPSS\WriteLog.txt");
        //2017/11/16 Add by DaiJun SPCで、判定画像検索画面で基板シリアルで検索をする際、入力をしても反応が無い場合があり、画像一覧も表示されない場合がある（No.178）[#7788 ] END

        #endregion

        #region コンストラクタ&デストラクタメ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public TopssLogger(string strfileRootPath, string strFileTitle)
        {
            lock (_lockObject)
            {
                _fileRootPath = strfileRootPath;
                _fileTitle = strFileTitle;
                _filePath = GetFilePath();

                List<string> listFile = SearchLogFile(GetFilePath(), _fileTitle);
                if (listFile.Count > 0)
                {
                    OpenFile(listFile.Last());
                }
                else
                {
                    OpenFile("");
                }

            }
        }

        /// <summary>
        /// デストラクタメ
        /// </summary>
        ~TopssLogger()
        {
            Dispose(false);
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Dispose
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                lock (_lockObject)
                {
                    _Files.Dispose();
                }
            }
            _disposed = true;
        }

        #endregion

        #region ファイル開く＆閉じる

        /// <summary>
        /// バッファエリアを更新(HDDに保存)
        /// </summary>
        public void Flush()
        {
            try
            {
                lock (_lockObject)
                {
                    _Files.Flush();
                }

            }
            catch (Exception e)
            {
                Debug.WriteLine("Flush" + e.Message);
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// ログレベルをチェック
        /// </summary>
        /// <param name="logLevel">ログレベル</param>
        /// <returns></returns>
        private bool IsWriteLogCheck(LogLevel logLevel)
        {
            //2017/11/16 Add by DaiJun SPCで、判定画像検索画面で基板シリアルで検索をする際、入力をしても反応が無い場合があり、画像一覧も表示されない場合がある（No.178））[#7788 ] START
            if (System.IO.File.Exists(pathlog) == false)
                return false;
            //2017/11/16 Add by DaiJun SPCで、判定画像検索画面で基板シリアルで検索をする際、入力をしても反応が無い場合があり、画像一覧も表示されない場合がある（No.178）[#7788 ] END

            if ((_logOutputLevels & logLevel) > 0)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// ログファイルを開く
        /// </summary>
        private void OpenFile(string fileName)
        {
            try
            {

                //Path名を判断
                _filePath = GetFilePath();

                //CHN 2017/10/20 Delete by DaiJun 去掉每天保留Log文件个数上限（LogFileNumPerDayLimit）Start

                //List<string> listFile = SearchLogFile(_filePath, _fileTitle);
                //while (listFile.Count >= LogFileNumPerDayLimit)
                //{
                //    if (listFile.Count > 1)
                //    {
                //        File.Delete(listFile[0]);
                //        listFile.RemoveAt(0);
                //    }
                //}

                //CHN 2017/10/20 Delete by DaiJun 去掉每天保留Log文件个数上限（LogFileNumPerDayLimit）End

                //ファイル名
                if (string.IsNullOrEmpty(fileName))
                {
                    _fileName = GetFileName();
                }
                else
                {
                    _fileName = fileName;
                }

                //ファイルフールパス
                _fileFullPath = Path.Combine(_filePath, _fileName);
                FileInfo fi = new FileInfo(_fileFullPath);
                if (!fi.Exists)
                {
                    lock (_lockObject)
                    {
                        _Files = fi.CreateText();
                    }
                }
                else
                {
                    lock (_lockObject)
                    {
                        _Files = fi.AppendText();
                    }
                }

            }
            catch (Exception e)
            {
                Debug.WriteLine("OpenFile : " + e.Message);
            }
        }

        /// <summary>
        /// 閉じる
        /// </summary>
        public void CloseFile()
        {
            try
            {
                lock (_lockObject)
                {
                    _Files.Close();
                    _Files.Dispose();
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("CloseFile : " + e.Message);
            }
        }

        /// <summary>
        /// 新しいファイルを作成する必要があるかどうか
        /// </summary>
        /// <returns></returns>
        private bool CheckCreateNewFile()
        {
            bool bChangeNewFile = false;
            try
            {
                //Path名を判断
                if (_filePath != GetFilePath())
                {
                    bChangeNewFile = true;
                }

                //ファイルサイズを判断
                FileInfo info = new FileInfo(_fileFullPath);
                if (info.Exists && info.Length / 1024 / 1024 >= LogFileSizeLimit)
                {
                    bChangeNewFile = true;
                }
                info = null;
            }
            catch (Exception e)
            {
                Debug.WriteLine("CheckCreateNewFile : " + e.Message);
            }
            return bChangeNewFile;
        }

        /// <summary>
        /// 出力ファイル名を取得
        /// </summary>
        /// <returns></returns>
        private string GetFileName()
        {
            string strFileName = DateTime.Now.ToString("yyyyMMddHHmmss");
            if (!string.IsNullOrEmpty(_fileTitle))
            {
                strFileName = _fileTitle + "_" + strFileName;
            }
            int n = 0;
            while (true)
            {
                if (!File.Exists(Path.Combine(_filePath, strFileName))) break;
                strFileName += "_" + n;
            }

            return strFileName + LogFileExtension;
        }

        /// <summary>
        /// 出力パスを取得
        /// </summary>
        /// <returns></returns>
        private string GetFilePath()
        {
            string strFilePath = "";
            if (string.IsNullOrEmpty(_fileRootPath))
            {
                _fileRootPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            }

            strFilePath = Path.Combine(_fileRootPath, DateTime.Now.ToString("yyyyMMdd"));
            if (!Directory.Exists(strFilePath))
            {
                Directory.CreateDirectory(strFilePath);
            }
            return strFilePath;
        }

        /// <summary>
        /// ログを書き込む
        /// </summary>
        private void Writeln(string strLine)
        {
            try
            {
                lock (_lockObject)
                {
                    //ファイルを切り替え判断
                    if (CheckCreateNewFile())
                    {
                        CloseFile();
                        OpenFile("");
                    }
                    //ファイル書き込む
                    _Files.WriteLine(strLine);
                }
                //更新
                Flush();
            }
            catch (Exception e)
            {
                Debug.WriteLine("Writeln : " + e.Message);
            }
        }

        private List<string> SearchLogFile(string filePath, string strTitleName)
        {
            List<string> listFilePath = new List<string>();
            Stopwatch st = new Stopwatch();
            st.Start();
            DirectoryInfo dir = new DirectoryInfo(filePath);
            if (!string.IsNullOrEmpty(strTitleName))
            {
                strTitleName += "_";
            }
            string strRegularExpression = @"^" + strTitleName + @"\d{4}[0-1][0-9][0-3][0-9][0-2][0-9][0-6][0-9][0-6][0-9].*\.log$";

            foreach (FileInfo fileInfo in dir.GetFiles())
            {

                if (System.Text.RegularExpressions.Regex.Matches(fileInfo.Name, strRegularExpression,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count > 0)
                {
                    listFilePath.Add(fileInfo.FullName);
                }
            }
            listFilePath.Sort((p1, p2) => String.Compare(p1, p2, StringComparison.Ordinal));
            st.Stop();
            Debug.WriteLine(st.ElapsedMilliseconds);
            return listFilePath;
        }

        #endregion

        #region 公開メソッド

        /// <summary>
        /// ログを書き込む
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="strMessage"></param>
        /// <param name="e"></param>
        public void Write(LogLevel logType, string strMessage, Exception e)
        {
            if (!IsWriteLogCheck(logType))
            {
                return;
            }
            
            string strLog = LogFormater.FormatLog(logType, strMessage, e);

            Writeln(strLog);
        }

        /// <summary>
        /// ログを書き込む
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="strMessage"></param>
        public void Write(LogLevel logType, string strMessage)
        {
            if (!IsWriteLogCheck(logType))
            {
                return;
            }

            string strLog = LogFormater.FormatLog(logType, strMessage, null);
            Writeln(strLog);
        }

        /// <summary>
        /// メソッド開始
        /// </summary>
        /// <param name="strMessage"></param>
        public void WriteMethodStart(string strMessage)
        {

            if (!IsWriteLogCheck(LogLevel.Info))
            {
                return;
            }

            string strLog = LogFormater.FormatLog(LogLevel.Info, strMessage + " START", null);
            Writeln(strLog);
        }

        /// <summary>
        /// メソッド終了
        /// </summary>
        /// <param name="strMessage"></param>
        public void WriteMethodEnd(string strMessage)
        {
            if (!IsWriteLogCheck(LogLevel.Info))
            {
                return;
            }

            string strLog = LogFormater.FormatLog(LogLevel.Info, strMessage + " END", null);
            Writeln(strLog);
        }

        /// <summary>
        /// ログラベルを設定
        /// </summary>
        /// <param name="nLogLevel"></param>
        /// <param name="bOutput"></param>
        public void SetLogOutputLevel(LogLevel nLogLevel, bool bOutput)
        {
            if (bOutput)
            {
                _logOutputLevels |= nLogLevel;
            }
            else
            {
                _logOutputLevels &= ~nLogLevel;
            }
        }

        #endregion
    }
}
