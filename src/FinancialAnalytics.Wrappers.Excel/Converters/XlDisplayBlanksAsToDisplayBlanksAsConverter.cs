using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public static class XlDisplayBlanksAsToDisplayBlanksAsConverter
    {
        public static DisplayBlanksAs ConvertBack(XlDisplayBlanksAs xlDisplayBlanksAs)
        {
            DisplayBlanksAs displayBlanksAs;
            switch (xlDisplayBlanksAs)
            {
                case XlDisplayBlanksAs.xlInterpolated:
                    displayBlanksAs = DisplayBlanksAs.Interpolated;
                    break;
                case XlDisplayBlanksAs.xlNotPlotted:
                    displayBlanksAs = DisplayBlanksAs.NotPlotted;
                    break;
                case XlDisplayBlanksAs.xlZero:
                    displayBlanksAs = DisplayBlanksAs.Zero;
                    break;
				default:
					displayBlanksAs = DisplayBlanksAs.NotPlotted;
                    break;
            }
            return displayBlanksAs;
        }

        public static XlDisplayBlanksAs ConvertBack(DisplayBlanksAs displayBlanksAs)
        {
            XlDisplayBlanksAs xlDisplayBlanksAs;
            switch (displayBlanksAs)
            {
                case DisplayBlanksAs.Interpolated:
                    xlDisplayBlanksAs = XlDisplayBlanksAs.xlInterpolated;
                    break;
                case DisplayBlanksAs.NotPlotted:
                    xlDisplayBlanksAs = XlDisplayBlanksAs.xlNotPlotted;
                    break;
                case DisplayBlanksAs.Zero:
                    xlDisplayBlanksAs = XlDisplayBlanksAs.xlZero;
                    break;
                default:
                    xlDisplayBlanksAs = XlDisplayBlanksAs.xlNotPlotted;
                    break;
            }
            return xlDisplayBlanksAs;
        }
    }
}
