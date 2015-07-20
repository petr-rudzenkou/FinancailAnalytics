using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlColorIndexToColorIndexConverter
    {
        public static ColorIndex Convert(XlColorIndex xlColorIndex)
        {
            switch (xlColorIndex)
            { 
                case XlColorIndex.xlColorIndexAutomatic:
                    return ColorIndex.ColorIndexAutomatic;
                case XlColorIndex.xlColorIndexNone:
                    return ColorIndex.ColorIndexNone;
                default:
                    return ColorIndex.ColorIndexAutomatic;
            }
        }

        public static XlColorIndex ConvertBack(ColorIndex colorIndex)
        {
            switch (colorIndex)
            {
                case ColorIndex.ColorIndexNone:
                    return XlColorIndex.xlColorIndexNone;
                case ColorIndex.ColorIndexAutomatic:
                default:
                    return XlColorIndex.xlColorIndexAutomatic;
            }
        }
    }
}
