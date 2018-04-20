using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace TopssLogger
{
    /// <summary>
    /// ログ出力レベル
    /// </summary>
    [Flags]
    public enum LogLevel
    {
        /// <summary>
        /// なし
        /// </summary>
        None = 0x00,
        /// <summary>
        /// Debug
        /// </summary>
        Debug = 0x01,
        /// <summary>
        /// Info
        /// </summary>
        Info = 0x02,
        /// <summary>
        /// Warning
        /// </summary>
        Warn = 0x04,
        /// <summary>
        /// Error
        /// </summary>
        Error = 0x08,
        /// <summary>
        /// Fatal
        /// </summary>
        Fatal = 0x10,
        /// <summary>
        /// すべて
        /// </summary>
        All = 0x1F
    }
}
