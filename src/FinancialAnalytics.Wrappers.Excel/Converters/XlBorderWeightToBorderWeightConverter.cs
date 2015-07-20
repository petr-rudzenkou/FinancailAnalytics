using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlBorderWeightToBorderWeightConverter
    {
        public static BorderWeight Convert(XlBorderWeight xlBorderWeight)
        {
            BorderWeight borderWeight;
            switch (xlBorderWeight)
            {
                case XlBorderWeight.xlHairline:
                    borderWeight = BorderWeight.Hairline;
                    break;
                case XlBorderWeight.xlMedium:
                    borderWeight = BorderWeight.Medium;
                    break;
                case XlBorderWeight.xlThick:
                    borderWeight = BorderWeight.Thick;
                    break;
                default:
                    borderWeight = BorderWeight.Thin;
                    break;
            }
            return borderWeight;
        }

        public static XlBorderWeight ConvertBack(BorderWeight borderWeight)
        {
            XlBorderWeight xlBorderWeight;
            switch (borderWeight)
            {
                case BorderWeight.Hairline:
                    xlBorderWeight = XlBorderWeight.xlHairline;
                    break;
                case BorderWeight.Medium:
                    xlBorderWeight = XlBorderWeight.xlMedium;
                    break;
                case BorderWeight.Thick:
                    xlBorderWeight = XlBorderWeight.xlThick;
                    break;
                default:
                    xlBorderWeight = XlBorderWeight.xlThin;
                    break;
            }
            return xlBorderWeight;
        }
    }
}
