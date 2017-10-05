using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordToTFS
{
    public class ClipboardHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("user32.dll")]
        private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")]
        private static extern bool CloseClipboard();
        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll")]
        private static extern uint GlobalSize(IntPtr hMem);
        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalUnlock(IntPtr hMem);
        [DllImport("user32.dll")]
        private static extern bool EmptyClipboard();
        [DllImport("user32.dll")]
        private static extern uint RegisterClipboardFormat(string lpszFormat);

        private ClipboardHelper()
        {
        }

        private static ClipboardHelper clipboardHelper;

        public static ClipboardHelper Instanse
        {
            get
            {
                if (clipboardHelper == null)
                      clipboardHelper = new ClipboardHelper();

                return clipboardHelper;
            }
        }

        /// <summary>
        /// Copy to clipboard
        /// </summary>
        /// <param name="text">text</param>
        public void CopyToClipboard(string text)
        {
            IntPtr pHtml = IntPtr.Zero;
            try
            {
                if (!OpenClipboard(IntPtr.Zero))
                    throw new Exception("Failed to open clipboard");
                EmptyClipboard();
                byte[] bytes = Encoding.UTF8.GetBytes(text);
                pHtml = Marshal.AllocHGlobal(bytes.Length);
                
                IntPtr pMFP = GlobalLock(pHtml);
                Marshal.Copy(bytes, 0, pMFP, bytes.Length);
                SetClipboardData(RegisterClipboardFormat(DataFormats.Html), pMFP); 
            }
            catch { }
            finally
            {
                CloseClipboard();
                if (pHtml != IntPtr.Zero)
                    GlobalUnlock(pHtml);
            }
        }

        /// <summary>
        /// Copy from clipboard
        /// </summary>
        public string CopyFromClipboard()
        {
            string text = string.Empty;
            IntPtr hGMem = IntPtr.Zero;

            try
            {
                if (!OpenClipboard(IntPtr.Zero))
                    throw new Exception("Failed to open clipboard");

                hGMem = GetClipboardData(RegisterClipboardFormat(DataFormats.Html));
                IntPtr pMFP = GlobalLock(hGMem);
                uint len = GlobalSize(hGMem);
                byte[] bytes = new byte[len];
                Marshal.Copy(pMFP, bytes, 0, (int)len);
                text = System.Text.Encoding.UTF8.GetString(bytes);  
            }
            catch { }
            finally
            {
                CloseClipboard();
                if (hGMem != IntPtr.Zero)
                    GlobalUnlock(hGMem);
            }
           
            return text;
        }
    }
}
