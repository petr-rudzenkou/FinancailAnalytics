using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlAxisGroupToAxisGroupConverter
    {
        public static AxisGroup Convert(XlAxisGroup xlAxisGroup)
        {
            AxisGroup axisGroup;
            switch (xlAxisGroup)
            {
                case XlAxisGroup.xlPrimary:
                    axisGroup = AxisGroup.Primary;
                    break;
                case XlAxisGroup.xlSecondary:
                    axisGroup = AxisGroup.Secondary;
                    break;
                default:
                    axisGroup = AxisGroup.Secondary;
                    break;
            }
            return axisGroup;
        }

        public static XlAxisGroup ConvertBack(AxisGroup axisGroup)
        {
            XlAxisGroup xlAxisGroup;
            switch (axisGroup)
            {
                case AxisGroup.Primary:
                    xlAxisGroup = XlAxisGroup.xlPrimary;
                    break;
                case AxisGroup.Secondary:
                    xlAxisGroup = XlAxisGroup.xlSecondary;
                    break;
                default:
                    xlAxisGroup = XlAxisGroup.xlSecondary;
                    break;
            }
            return xlAxisGroup;            
        }
    }
}
