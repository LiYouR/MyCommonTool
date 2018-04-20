using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace CommonTool
{
    public class DirectoryEx
    {
        public static void DeleteDirectory(string target_dir)
        {
            string[] files = System.IO.Directory.GetFiles(target_dir);
            string[] dirs = System.IO.Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

            string[] dirs2 = System.IO.Directory.GetDirectories(target_dir);
            int nWaitTimes = 0;
            while (dirs2.Length > 0)
            {
                Thread.Sleep(5);
                nWaitTimes++;
                if (nWaitTimes > 3)
                {
                    break;
                }
            }
            System.IO.Directory.Delete(target_dir, false);
        }

		//2018/02/06 ADD BY LIYOU [#741] [TOPSS]ログ一括収集(#7817) START
        public static void TraversalDirectory(string strPath, Action<string> action, bool recursive = true)
        {
            //path exist?
            if (!Directory.Exists(strPath))
            {
                return;
            }

            DirectoryInfo dir = new DirectoryInfo(strPath);
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)
                {
                    if (recursive)
                    {
                        TraversalDirectory(fsinfo.FullName, action);  
                    }
                }
                else
                {
                    action(fsinfo.FullName);
                }
            }
        }

        public static void DeleteEmptyDirectory(string strPath)
        {
            //path exist?
            if (!Directory.Exists(strPath))
            {
                return;
            }

            DirectoryInfo dir = new DirectoryInfo(strPath);
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            if (fsinfos.Length == 0)
            {
                Directory.Delete(strPath);
                return;
            }
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)
                {
                    DeleteEmptyDirectory(fsinfo.FullName);
                }
            }
        }
		//2018/02/06 ADD BY LIYOU [#741] [TOPSS]ログ一括収集(#7817) END
    }
}
