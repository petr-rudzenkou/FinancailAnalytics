using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlAxisTypeToAxisTypeConverter
    {
        public static AxisType Convert(XlAxisType xlAxisType)
        {
            AxisType axisType;
            switch (xlAxisType)
            {
                case XlAxisType.xlCategory:
                    axisType = AxisType.Category;
                    break;
                case XlAxisType.xlSeriesAxis:
                    axisType = AxisType.SeriesAxis;
                    break;
                default:
                    axisType = AxisType.Value;
                    break;
            }
            return axisType;
        }

        public static XlAxisType ConvertBack(AxisType axisType)
        {
            XlAxisType xlAxisType;
            switch (axisType)
            {
                case AxisType.Category:
                    xlAxisType = XlAxisType.xlCategory;
                    break;
                case AxisType.SeriesAxis:
                    xlAxisType = XlAxisType.xlSeriesAxis;
                    break;
                default:
                    xlAxisType = XlAxisType.xlValue;
                    break;
            }
            return xlAxisType;
        }
    }
}
