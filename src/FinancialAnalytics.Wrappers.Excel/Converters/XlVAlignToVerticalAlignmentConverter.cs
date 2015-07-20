using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlVAlignToVerticalAlignmentConverter
    {
        public static VerticalAlignment Convert(XlVAlign xlVAlign)
        {
            VerticalAlignment verticalAlignment;
            switch (xlVAlign)
            {
                case XlVAlign.xlVAlignBottom:
                    verticalAlignment = VerticalAlignment.Bottom;
                    break;
                case XlVAlign.xlVAlignCenter:
                    verticalAlignment = VerticalAlignment.Center;
                    break;
                case XlVAlign.xlVAlignDistributed:
                    verticalAlignment = VerticalAlignment.Distributed;
                    break;
                case XlVAlign.xlVAlignJustify:
                    verticalAlignment = VerticalAlignment.Justify;
                    break;
                default:
                    verticalAlignment = VerticalAlignment.Top;
                    break;
            }
            return verticalAlignment;
        }

        public static XlVAlign ConvertBack(VerticalAlignment verticalAlignment)
        {
            XlVAlign xlVAlign;
            switch (verticalAlignment)
            {
                case VerticalAlignment.Bottom:
                    xlVAlign = XlVAlign.xlVAlignBottom;
                    break;
                case VerticalAlignment.Center:
                    xlVAlign = XlVAlign.xlVAlignCenter;
                    break;
                case VerticalAlignment.Distributed:
                    xlVAlign = XlVAlign.xlVAlignDistributed;
                    break;
                case VerticalAlignment.Justify:
                    xlVAlign = XlVAlign.xlVAlignJustify;
                    break;
                default:
                    xlVAlign = XlVAlign.xlVAlignTop;
                    break;
            }
            return xlVAlign;
        }
    }
}
