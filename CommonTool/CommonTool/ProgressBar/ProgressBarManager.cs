using System;
using System.ComponentModel;

namespace CommonTool
{
    public class ProgressBarManager
    {
        public enum EResult
        {
            OK = 0,
            Failed = -1,
            Cannceled = -2
        }

        //Singleton
        private static volatile ProgressBarManager _instance;
        private static readonly object _lock = new object();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        private ProgressBarManager()
        {
            _backgroundWorker = null;
        }

        /// <summary>
        /// インスタンス取得
        /// </summary>
        /// <returns></returns>
        public static ProgressBarManager GetInstance()
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                        _instance = new ProgressBarManager();
                }
            }
            return _instance;
        }

        //ワークスレッド
        private BackgroundWorker _backgroundWorker;

        //Progressダイアログ
        private IProgressForm _processForm;

        //実行方法
        private Func<object, int> _work;

        //キャンセルイベント
        private Action _cancel; 

        //パラメーター
        private object objectParameter;

        //実行結果
        private int _workResult;

        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="ProgressForm">Progressダイアログ</param>
        /// <param name="work">実行方法</param>
        /// <param name="o">パラメーター</param>
        /// <returns></returns>
        public bool InitBackgroundWorker(IProgressForm ProgressForm, Func<object, int> work, object o, Action cancel)
        {
            if (null != _backgroundWorker)
            {
                return false;
            }
            _workResult = (int)EResult.Failed;
            _work = work;
            _cancel = cancel;
            objectParameter = o;
            _processForm = ProgressForm;
            _processForm.StopEvent_IProgress += Stop;
            _processForm.StartTask_IProgress += StartTask;

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.WorkerSupportsCancellation = true;
            _backgroundWorker.DoWork += BackgroundWorker_DoWork;
            _backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            _backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            return true;
        }

        /// <summary>
        /// ProgressChangedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _processForm.ProgressValue_IProgress = e.ProgressPercentage;
        }

        /// <summary>
        /// Completedイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_processForm.Visible_IProgress)
            {
                _processForm.Close_IProgress();
            }

            _backgroundWorker.Dispose();
            _backgroundWorker = null;
            _work = null;
            objectParameter = null;
            _processForm = null;
        }

        /// <summary>
        /// DoWorkイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_work != null)
            {
                _workResult = _work(objectParameter);
            }
        }

        /// <summary>
        /// 中止
        /// </summary>
        private void Stop()
        {
            _backgroundWorker.CancelAsync();
            if (null != _cancel)
            {
                _cancel();
            }
        }

        /// <summary>
        /// 開始
        /// </summary>
        /// <returns></returns>
        public bool Start()
        {
            if (null == _backgroundWorker)
            {
                return false;
            }
            _processForm.ShowDialog_IProgress();
            //_backgroundWorker.RunWorkerAsync();
            return true;
        }

        private void StartTask()
        {
            _backgroundWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Progress設定
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public bool ReportProgress(int i)
        {
            if (null == _backgroundWorker)
            {
                return false;
            }
            _backgroundWorker.ReportProgress(i);
            return true;
        }

        /// <summary>
        /// 中止請求があるかどうか
        /// </summary>
        /// <returns></returns>
        public bool IsCancellationPending()
        {
            if (null == _backgroundWorker)
            {
                return false;
            }
            return _backgroundWorker.CancellationPending;
        }

        /// <summary>
        /// 実行結果を取得
        /// </summary>
        /// <returns></returns>
        public int GetResult()
        {
            return _workResult;
        }
    }
}
