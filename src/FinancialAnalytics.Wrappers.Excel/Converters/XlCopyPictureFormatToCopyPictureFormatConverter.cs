using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlCopyPictureFormatToCopyPictureFormatConverter
    {
        public static CopyPictureFormat Convert(XlCopyPictureFormat xlCopyPictureFormat)
        {
            CopyPictureFormat copyPictureFormat;
            switch (xlCopyPictureFormat)
            {
                case XlCopyPictureFormat.xlBitmap:
                    copyPictureFormat = CopyPictureFormat.Bitmap;
                    break;
                default:
                    copyPictureFormat = CopyPictureFormat.Picture;
                    break;
            }
            return copyPictureFormat;
        }

        public static XlCopyPictureFormat ConvertBack(CopyPictureFormat copyPictureFormat)
        {
            XlCopyPictureFormat xlCopyPictureFormat;
            switch (copyPictureFormat)
            {
                case CopyPictureFormat.Bitmap:
                    xlCopyPictureFormat = XlCopyPictureFormat.xlBitmap;
                    break;
                default:
                    xlCopyPictureFormat = XlCopyPictureFormat.xlPicture;
                    break;
            }
            return xlCopyPictureFormat;
        }
    }
}
