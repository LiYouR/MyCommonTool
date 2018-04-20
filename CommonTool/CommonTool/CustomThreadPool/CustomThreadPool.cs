using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;


namespace CommonTool
{
    public class CustomThreadPool : IDisposable
    {
        // 実行中エラー発生デリゲート
        public delegate void RunErrorEventHandler(string strErrorMsg);

        /// 実行中エラー発生イベントハンドラ
        public event RunErrorEventHandler RunError;
        protected virtual void OnRunError(string strErrMsg)
        {
            if (RunError != null) RunError(strErrMsg);
        }

        //シグナル
        public static Semaphore semaphore;

        //スレッド数
        private int _threadNum;

        //キャンセル事件
        private CancellationTokenSource _cancelTokenSource;

        //タスクトキュー
        private ConcurrentQueue<ICustomTask> _taskQueue;

        //スレッドリスト
        private List<Thread> _workThreadList;

        //ClosedFlag
        private bool bClosed;
        private object _lock;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CustomThreadPool(int i)
        {
            bClosed = false;
            _lock = new object();
            semaphore = new Semaphore(0, Int32.MaxValue);
            _taskQueue = new ConcurrentQueue<ICustomTask>();
            _cancelTokenSource = new CancellationTokenSource();
            _threadNum = Environment.ProcessorCount - 1;
            if (i > 0)
            {
                _threadNum = i;
            }
            CreateThreadPool();
        }

        /// <summary>
        /// コンストラクタ（スレッド数 =　Processor - 1）
        /// </summary>
        public CustomThreadPool()
            : this(0)
        {
        }

        /// <summary>
        /// ThreadPool初期化
        /// </summary>
        private void CreateThreadPool()
        {
            if (_workThreadList == null)
                _workThreadList = new List<Thread>();

            lock (_workThreadList)
            {
                for (int i = 0; i < _threadNum; i++)
                {
                    Thread thread = new Thread(ThreadRunProc);
                    _workThreadList.Add(thread);
                    thread.Start();
                }
            }
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            CloseThreadPool();
            _cancelTokenSource.Dispose();
            semaphore.Close();
            _taskQueue = null;
            _workThreadList = null;
        }

        /// <summary>
        /// タスク追加
        /// </summary>
        /// <param name="task"></param>
        public void AddTask(ICustomTask task)
        {
            if (task == null)
                return;
            lock (_taskQueue)
            {
                task.cancelTokenSource = _cancelTokenSource;
                _taskQueue.Enqueue(task);
                semaphore.Release();
            }

        }

        /// <summary>
        /// 処理中止
        /// </summary>
        public void Cancel()
        {
            CloseThreadPool();
        }

        /// <summary>
        /// スレッド数を取得
        /// </summary>
        public int GetThreadCount()
        {
            return _threadNum;
        }

        /// <summary>
        /// ThreadPool閉じる
        /// </summary>
        private void CloseThreadPool()
        {
            if (bClosed) return;
            lock (_lock)
            {
                if (bClosed) return;

                _cancelTokenSource.Cancel();

                semaphore.Release(_workThreadList.Count);

                while (true)
                {
                    lock (_workThreadList)
                    {
                        for (int i = _workThreadList.Count - 1; i >= 0; i--)
                        {
                            Thread thread = _workThreadList[i];
                            if (null != thread && !thread.IsAlive)
                            {
                                _workThreadList.RemoveAt(i);
                            }
                        }
                    }
                    if (_workThreadList.Count > 0)
                    {
                        Thread.Sleep(50);
                    }
                    else
                    {
                        break;
                    }
                }
                bClosed = true;
            }
        }

        /// <summary>
        /// スレッドRunメソッド
        /// </summary>
        private void ThreadRunProc()
        {
            while (semaphore.WaitOne()
                && !_cancelTokenSource.IsCancellationRequested)
            {
                bool bHaveMoreTask;
                //タスク取得
                ICustomTask customTask;
                lock (_taskQueue)
                {
                    bHaveMoreTask = _taskQueue.TryDequeue(out customTask);
                }

                if (!bHaveMoreTask)
                {
                    semaphore.Release();
                    continue;
                }

                //キューにタスクあり
                if (null != customTask)
                {
                    //イベント実行
                    bool bRet = customTask.Excute();
                    if (!bRet)
                    {
                        //エラーの場合、エラー処理へ
                        OnRunError(customTask.GetLastError());
                    }
                }
            }
        }
    }
}
