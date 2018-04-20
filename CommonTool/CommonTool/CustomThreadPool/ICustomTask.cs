using System.Threading;

namespace CommonTool
{
    public interface ICustomTask
    {
        /// <summary>
        /// キャンセル事件
        /// </summary>
        CancellationTokenSource cancelTokenSource { set; }

        /// <summary>
        /// タスク実行
        /// </summary>
        bool Excute();

        /// <summary>
        /// エラーメッセージ取得
        /// </summary>
        string GetLastError();
    }
}
