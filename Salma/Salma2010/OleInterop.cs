using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Salma2010
{
    using WordToTFS;

    internal static class OleInterop
    {
        private const short PictureTypeBitmap = 1;

        private static ImageSource PictureDispToImage(this stdole.IPictureDisp pictureDisp)
        {
            BitmapSource image = null;
            if (pictureDisp != null && pictureDisp.Type == PictureTypeBitmap)
            {
                var paletteHandle = new IntPtr(pictureDisp.hPal);
                var bitmapHandle = new IntPtr(pictureDisp.Handle);
                image = Imaging.CreateBitmapSourceFromHBitmap(bitmapHandle, paletteHandle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            return image;
        }

        public static ImageSource GetMsoImage(Icons icon, int width = 16, int height = 16)
        {
            var msoId = OfficeHelper.GetImageMso(icon, Globals.ThisAddIn.MsWordVersion);
            return PictureDispToImage(Globals.ThisAddIn.Application.CommandBars.GetImageMso(msoId, width, height));
        }
    }
}

