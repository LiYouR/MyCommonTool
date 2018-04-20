using System;
using System.Drawing;
using System.IO;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CommonTool
{
    public class ViewCapture
    {
        [DllImport("user32.dll")]
        static extern void keybd_event
            (
            byte bVk,
            byte bScan,
            uint dwFlags,
            IntPtr dwExtraInfo
            );

        /// <summary>
        /// ALT + PRTSCNキーボートを押す 
        /// </summary>
        public void AltPrintScreen()
        {
            keybd_event((byte)Keys.Menu, 0, 0x0, IntPtr.Zero);
            keybd_event((byte)0x2c, 0, 0x0, IntPtr.Zero); //down
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();
            keybd_event((byte)0x2c, 0, 0x2, IntPtr.Zero); //up
            keybd_event((byte)Keys.Menu, 0, 0x2, IntPtr.Zero);
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// クリップボードに画像を取得する
        /// </summary>
        /// <returns></returns>
        public Bitmap GetScreenImage()
        {
            Bitmap bitMap = null;
            try
            {
                System.Windows.Forms.Application.DoEvents();
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    Image image = System.Windows.Forms.Clipboard.GetImage();
                    if (image != null)
                        bitMap = (Bitmap) (image.Clone());
                }
                else
                {

                }
                return bitMap;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }
}
