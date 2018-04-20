using System.IO;
using System.Threading;


namespace CommonTool.IOEx
{
    public class ZIP
    {

        enum SHFILEOPSTRUCT : int
        {
            /// <summary>
            /// Default. No options specified.
            /// </summary>
            FOF_DEFAULT = 0,

            /// <summary>
            /// Do not display a progress dialog box.
            /// </summary>
            FOF_SILENT = 4,

            /// <summary>
            /// Rename the target file if a file exists at the target location with the same name.
            /// </summary>
            FOF_RENAMEONCOLLISION = 8,

            /// <summary>
            /// Click "Yes to All" in any dialog box displayed.
            /// </summary>b
            FOF_NOCONFIRMATION = 16,

            /// <summary>
            /// Preserve undo information, if possible.
            /// </summary>
            FOF_ALLOWUNDO = 64,

            /// <summary>
            /// Perform the operation only if a wildcard file name (*.*) is specified.
            /// </summary>
            FOF_FILESONLY = 128,

            /// <summary>
            /// Display a progress dialog box but do not show the file names.
            /// </summary>
            FOF_SIMPLEPROGRESS = 256,

            /// <summary>
            /// Do not confirm the creation of a new directory if the operation requires one to be created.
            /// </summary>
            FOF_NOCONFIRMMKDIR = 512,

            /// <summary>
            /// Do not display a user interface if an error occurs.
            /// </summary>
            FOF_NOERRORUI = 1024,

            /// <summary>
            /// Do not copy the security attributes of the file
            /// </summary>
            FOF_NOCOPYSECURITYATTRIBS = 2048,

            /// <summary>
            /// Disable recursion.
            /// </summary>
            FOF_NORECURSION = 4096,

            /// <summary>
            /// Do not copy connected files as a group. Only copy the specified files.
            /// </summary>
            FOF_NO_CONNECTED_ELEMENTS = 9182
        }

        /// <summary>
        /// Zip function calls Shell32, Interop.Shell32.dll is needed
        /// </summary>
        /// <param name="filesInFolder">Specify a folder containing the zip source files</param>
        /// <param name="zipFile">Specify the final zip file name, with ".zip" extension</param>
        public static void Compress(string filesInFolder, string zipFile)
        {
            if (filesInFolder == null || filesInFolder.Trim() == "") return;
            if (zipFile == null || zipFile.Trim() == "") return;
            if (!Directory.Exists(filesInFolder)) return;

            DirectoryEx.DeleteEmptyDirectory(filesInFolder);

            Shell32.ShellClass sh = new Shell32.ShellClass();
            try
            {
                if (File.Exists(zipFile)) File.Delete(zipFile);
                //Create an empty zip file
                byte[] emptyzip = new byte[] {80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};

                FileStream fs = File.Create(zipFile);
                fs.Write(emptyzip, 0, emptyzip.Length);
                fs.Flush();
                fs.Close();

                Shell32.Folder srcFolder = sh.NameSpace(filesInFolder);
                Shell32.Folder destFolder = sh.NameSpace(zipFile);
                Shell32.FolderItems items = srcFolder.Items();
                destFolder.CopyHere(items, SHFILEOPSTRUCT.FOF_SILENT | SHFILEOPSTRUCT.FOF_NOCONFIRMATION);
                while(items.Count != destFolder.Items().Count) Thread.Sleep(50); 
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sh);
            }
        }
    }
}
