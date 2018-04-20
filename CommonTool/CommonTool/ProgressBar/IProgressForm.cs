using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CommonTool
{
    public delegate void DelegateProgressStop();
    public delegate void DelegateStartTask();

    public interface IProgressForm
    {
        //停止イベント
        event DelegateProgressStop StopEvent_IProgress;

        //タスク開始イベント
        event DelegateStartTask StartTask_IProgress;

        //Progress設定 1～100
        int ProgressValue_IProgress { set; }

        //Form Visible
        bool Visible_IProgress { get; }

        //Formを閉じる
        void Close_IProgress();

        //Formを開く
        void ShowDialog_IProgress();
    }
}
