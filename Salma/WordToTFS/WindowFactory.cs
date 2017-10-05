using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;

namespace WordToTFS
{
    public static class WindowFactory
    {
        public static Func<Icons, int, int, ImageSource> IconConverter;


        public static void Create<T>(this T control, string title, Icons icon) where T : Window, new()
        {

            Window win;

            if (control == null)
            {
                win = new T();
            }
            else
            {
                win = (T)control;
            }

            win.ResizeMode = ResizeMode.NoResize;
            win.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            win.ShowInTaskbar = false;
            win.Icon = IconConverter(icon, 32, 32);
            win.Title = title ?? win.Title;

            IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
            WindowInteropHelper helper = new WindowInteropHelper(win);
            helper.Owner = mainWindowHandle;

            Action action = win.Close;

            win.ShowDialog();
        }

       
    }
}