using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlScaleTypeToScaleTypeConverter
    {
        public static ScaleType Convert(XlScaleType xlScaleType)
        {
            ScaleType scaleType;
            switch (xlScaleType)
            {
                case XlScaleType.xlScaleLinear:
                    scaleType = ScaleType.ScaleLinear;
                    break;
                case XlScaleType.xlScaleLogarithmic:
                    scaleType = ScaleType.ScaleLogarithmic;
                    break;
                default:
                    scaleType = ScaleType.ScaleLogarithmic;
                    break;
            }
            return scaleType;
        }

        public static XlScaleType ConvertBack(ScaleType scaleType)
        {
            XlScaleType xlScaleType;
            switch (scaleType)
            {
                case ScaleType.ScaleLinear:
                    xlScaleType = XlScaleType.xlScaleLinear;
                    break;
                case ScaleType.ScaleLogarithmic:
                    xlScaleType = XlScaleType.xlScaleLogarithmic;
                    break;
                default:
                    xlScaleType = XlScaleType.xlScaleLogarithmic;
                    break;
            }
            return xlScaleType;
        }
    }
}
