using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlHAlignToHorizontalAlignmentConverter
    {
        public static HorizontalAlignment Convert(XlHAlign xlHAlign)
        {
            HorizontalAlignment horizontalAlignment;
            switch (xlHAlign)
            {
                case XlHAlign.xlHAlignCenter:
                    horizontalAlignment = HorizontalAlignment.Center;
                    break;
                case XlHAlign.xlHAlignCenterAcrossSelection:
                    horizontalAlignment = HorizontalAlignment.CenterAcrossSelection;
                    break;
                case XlHAlign.xlHAlignDistributed:
                    horizontalAlignment = HorizontalAlignment.Distributed;
                    break;
                case XlHAlign.xlHAlignFill:
                    horizontalAlignment = HorizontalAlignment.Fill;
                    break;
                case XlHAlign.xlHAlignJustify:
                    horizontalAlignment = HorizontalAlignment.Justify;
                    break;
                case XlHAlign.xlHAlignLeft:
                    horizontalAlignment = HorizontalAlignment.Left;
                    break;
                case XlHAlign.xlHAlignRight:
                    horizontalAlignment = HorizontalAlignment.Right;
                    break;
                default:
                    horizontalAlignment = HorizontalAlignment.General;
                    break;
            }
            return horizontalAlignment;
        }

        public static XlHAlign ConvertBack(HorizontalAlignment horizontalAlignment)
        {
            XlHAlign xlHAlign;
            switch (horizontalAlignment)
            {
                case HorizontalAlignment.Center:
                    xlHAlign = XlHAlign.xlHAlignCenter;
                    break;
                case HorizontalAlignment.CenterAcrossSelection:
                    xlHAlign = XlHAlign.xlHAlignCenterAcrossSelection;
                    break;
                case HorizontalAlignment.Distributed:
                    xlHAlign = XlHAlign.xlHAlignDistributed;
                    break;
                case HorizontalAlignment.Fill:
                    xlHAlign = XlHAlign.xlHAlignFill;
                    break;
                case HorizontalAlignment.Justify:
                    xlHAlign = XlHAlign.xlHAlignJustify;
                    break;
                case HorizontalAlignment.Left:
                    xlHAlign = XlHAlign.xlHAlignLeft;
                    break;
                case HorizontalAlignment.Right:
                    xlHAlign = XlHAlign.xlHAlignRight;
                    break;
                default:
                    xlHAlign = XlHAlign.xlHAlignGeneral;
                    break;
            }
            return xlHAlign;
        }
    }
}
