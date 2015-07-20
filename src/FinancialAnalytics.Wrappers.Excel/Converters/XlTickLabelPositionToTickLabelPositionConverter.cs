using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlTickLabelPositionToTickLabelPositionConverter
    {
        public static TickLabelPosition Convert(XlTickLabelPosition xlTickLabelPosition)
        {
            TickLabelPosition tickLabelPosition;
            switch (xlTickLabelPosition)
            {
                case XlTickLabelPosition.xlTickLabelPositionHigh:
                    tickLabelPosition = TickLabelPosition.High;
                    break;
                case XlTickLabelPosition.xlTickLabelPositionLow:
                    tickLabelPosition = TickLabelPosition.Low;
                    break;
                case XlTickLabelPosition.xlTickLabelPositionNextToAxis:
                    tickLabelPosition = TickLabelPosition.NextToAxis;
                    break;
                default:
                    tickLabelPosition = TickLabelPosition.None;
                    break;
            }
            return tickLabelPosition;
        }

        public static XlTickLabelPosition ConvertBack(TickLabelPosition tickLabelPosition)
        {
            XlTickLabelPosition xlTickLabelPosition;
            switch (tickLabelPosition)
            {
                case TickLabelPosition.High:
                    xlTickLabelPosition = XlTickLabelPosition.xlTickLabelPositionHigh;
                    break;
                case TickLabelPosition.Low:
                    xlTickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow;
                    break;
                case TickLabelPosition.NextToAxis:
                    xlTickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNextToAxis;
                    break;
                default:
                    xlTickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone;
                    break;
            }
            return xlTickLabelPosition;
        }
    }
}
