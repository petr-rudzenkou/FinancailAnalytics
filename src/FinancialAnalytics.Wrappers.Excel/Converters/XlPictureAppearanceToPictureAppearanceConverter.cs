using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPictureAppearanceToPictureAppearanceConverter
    {
        public static PictureAppearance Convert(XlPictureAppearance xlPictureAppearance)
        {
            PictureAppearance pictureAppearance;
            switch (xlPictureAppearance)
            {
                case XlPictureAppearance.xlPrinter:
                    pictureAppearance = PictureAppearance.Printer;
                    break;
                default:
                    pictureAppearance = PictureAppearance.Screen;
                    break;
            }
            return pictureAppearance;
        }

        public static XlPictureAppearance ConvertBack(PictureAppearance pictureAppearance)
        {
            XlPictureAppearance xlPictureAppearance;
            switch (pictureAppearance)
            {
                case PictureAppearance.Printer:
                    xlPictureAppearance = XlPictureAppearance.xlPrinter;
                    break;
                default:
                    xlPictureAppearance = XlPictureAppearance.xlScreen;
                    break;
            }
            return xlPictureAppearance;
        }
    }
}
