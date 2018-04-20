using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace CommonSystem
{
    public static class ToolMiniDump
    {
        [DllImport("DbgHelp.dll")]
        private static extern Boolean MiniDumpWriteDump(
            IntPtr hProcess,
            Int32 processId,
            IntPtr fileHandle,
            MiniDumpType dumpType,
            ref MinidumpExceptionInfo excepInfo,
            IntPtr useInfo,
            IntPtr extInfo
        );

        [StructLayout(LayoutKind.Sequential, Pack = 4)] 
        struct MinidumpExceptionInfo
        {
            public Int32 ThreadId;
            public IntPtr ExceptionPointers;
            [MarshalAs(UnmanagedType.Bool)]
            public Boolean ClientPointers;
        }

        public static Boolean TryDump(String dmpPath, MiniDumpType dmpType)
        {
            //Dumpファイルを作成
            using (FileStream stream = new FileStream(dmpPath, FileMode.Create))
            {
                //プロセス情報を取得
                Process process = Process.GetCurrentProcess();

                // MINIDUMP_EXCEPTION_INFORMATION 情報の初期化
                MinidumpExceptionInfo mei = new MinidumpExceptionInfo();

                mei.ThreadId = Thread.CurrentThread.ManagedThreadId;
                mei.ExceptionPointers = Marshal.GetExceptionPointers();
                mei.ClientPointers = true;


                //Win32 APIを呼び出し
                Boolean res = MiniDumpWriteDump(
                                    process.Handle,
                                    process.Id,
                                    stream.SafeFileHandle.DangerousGetHandle(),
                                    dmpType,
                                    ref mei,
                                    IntPtr.Zero,
                                    IntPtr.Zero);

                //ストリームをクリア
                stream.Flush();
                stream.Close();

                return res;
            }
        }

        public enum MiniDumpType
        {
            None = 0x00010000,
            Normal = 0x00000000,
            WithDataSegs = 0x00000001,
            WithFullMemory = 0x00000002,
            WithHandleData = 0x00000004,
            FilterMemory = 0x00000008,
            ScanMemory = 0x00000010,
            WithUnloadedModules = 0x00000020,
            WithIndirectlyReferencedMemory = 0x00000040,
            FilterModulePaths = 0x00000080,
            WithProcessThreadData = 0x00000100,
            WithPrivateReadWriteMemory = 0x00000200,
            WithoutOptionalData = 0x00000400,
            WithFullMemoryInfo = 0x00000800,
            WithThreadInfo = 0x00001000,
            WithCodeSegs = 0x00002000
        }
    }
}