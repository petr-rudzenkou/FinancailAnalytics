using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlLineStyleToLineStyleConverter
    {
        public static LineStyle Convert(XlLineStyle xlLineStyle)
        {
            LineStyle lineStyle;
            switch (xlLineStyle)
            {
                case XlLineStyle.xlContinuous:
                    lineStyle = LineStyle.Continuous;
                    break;
                case XlLineStyle.xlDash:
                    lineStyle = LineStyle.Dash;
                    break;
                case XlLineStyle.xlDashDot:
                    lineStyle = LineStyle.DashDot;
                    break;
                case XlLineStyle.xlDashDotDot:
                    lineStyle = LineStyle.DashDotDot;
                    break;
                case XlLineStyle.xlDot:
                    lineStyle = LineStyle.Dot;
                    break;
                case XlLineStyle.xlDouble:
                    lineStyle = LineStyle.Double;
                    break;
                case XlLineStyle.xlLineStyleNone:
                    lineStyle = LineStyle.LineStyleNone;
                    break;
                default:
                    lineStyle = LineStyle.SlantDashDot;
                    break;
            }
            return lineStyle;
        }

        public static XlLineStyle ConvertBack(LineStyle lineStyle)
        {
            XlLineStyle xlLineStyle;
            switch (lineStyle)
            {
                case LineStyle.Continuous:
                    xlLineStyle = XlLineStyle.xlContinuous;
                    break;
                case LineStyle.Dash:
                    xlLineStyle = XlLineStyle.xlDash;
                    break;
                case LineStyle.DashDot:
                    xlLineStyle = XlLineStyle.xlDashDot;
                    break;
                case LineStyle.DashDotDot:
                    xlLineStyle = XlLineStyle.xlDashDotDot;
                    break;
                case LineStyle.Dot:
                    xlLineStyle = XlLineStyle.xlDot;
                    break;
                case LineStyle.Double:
                    xlLineStyle = XlLineStyle.xlDouble;
                    break;
                case LineStyle.LineStyleNone:
                    xlLineStyle = XlLineStyle.xlLineStyleNone;
                    break;
                default:
                    xlLineStyle = XlLineStyle.xlSlantDashDot;
                    break;
            }
            return xlLineStyle;
        }
    }
}
