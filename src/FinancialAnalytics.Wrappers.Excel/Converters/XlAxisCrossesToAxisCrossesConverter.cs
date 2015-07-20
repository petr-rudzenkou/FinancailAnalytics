using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlAxisCrossesToAxisCrossesConverter
    {
        public static AxisCrosses Convert(XlAxisCrosses xlAxisCrosses)
        {
            AxisCrosses axisCrosses;
            switch (xlAxisCrosses)
            {
                case XlAxisCrosses.xlAxisCrossesAutomatic:
                    axisCrosses = AxisCrosses.AxisCrossesAutomatic;
                    break;
                case XlAxisCrosses.xlAxisCrossesCustom:
                    axisCrosses = AxisCrosses.AxisCrossesCustom;
                    break;
                case XlAxisCrosses.xlAxisCrossesMaximum:
                    axisCrosses = AxisCrosses.AxisCrossesMaximum;
                    break;
                case XlAxisCrosses.xlAxisCrossesMinimum:
                    axisCrosses = AxisCrosses.AxisCrossesMinimum;
                    break;
                default:
                    axisCrosses = AxisCrosses.AxisCrossesAutomatic;
                    break;
            }
            return axisCrosses;
        }

        public static XlAxisCrosses Convert(AxisCrosses axisCrosses)
        {
            XlAxisCrosses xlAxisCrosses;
            switch (axisCrosses)
            {
                case AxisCrosses.AxisCrossesAutomatic:
                    xlAxisCrosses = XlAxisCrosses.xlAxisCrossesAutomatic;
                    break;
                case AxisCrosses.AxisCrossesCustom:
                    xlAxisCrosses = XlAxisCrosses.xlAxisCrossesCustom;
                    break;
                case AxisCrosses.AxisCrossesMaximum:
                    xlAxisCrosses = XlAxisCrosses.xlAxisCrossesMaximum;
                    break;
                case AxisCrosses.AxisCrossesMinimum:
                    xlAxisCrosses = XlAxisCrosses.xlAxisCrossesMinimum;
                    break;
                default:
                    xlAxisCrosses = XlAxisCrosses.xlAxisCrossesAutomatic;
                    break;
            }
            return xlAxisCrosses;
        }
    }
}
