using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlSheetVisibilityToSheetVisibilityConverter
    {
        public static SheetVisibility Convert(XlSheetVisibility xlSheetVisibility)
        {
            SheetVisibility sheetVisibility;
            switch (xlSheetVisibility)
            {
                case XlSheetVisibility.xlSheetHidden:
                    sheetVisibility = SheetVisibility.Hidden;
                    break;
                case XlSheetVisibility.xlSheetVeryHidden:
                    sheetVisibility = SheetVisibility.VeryHidden;
                    break;
                default:
                    sheetVisibility = SheetVisibility.Visible;
                    break;
            }
            return sheetVisibility;
        }

        public static XlSheetVisibility ConvertBack(SheetVisibility sheetVisibility)
        {
            XlSheetVisibility xlSheetVisibility;
            switch (sheetVisibility)
            {
                case SheetVisibility.Hidden:
                    xlSheetVisibility = XlSheetVisibility.xlSheetHidden;
                    break;
                case SheetVisibility.VeryHidden:
                    xlSheetVisibility = XlSheetVisibility.xlSheetVeryHidden;
                    break;
                default:
                    xlSheetVisibility = XlSheetVisibility.xlSheetVisible;
                    break;
            }
            return xlSheetVisibility;
        }
    }
}
