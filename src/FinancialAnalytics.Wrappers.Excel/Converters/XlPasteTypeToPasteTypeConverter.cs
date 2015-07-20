using System;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPasteTypeToPasteTypeConverter
    {
        public static PasteType Convert(XlPasteType xlPasteType)
        {
            PasteType chartType = PasteType.PasteAll;
            string chartTypeName = xlPasteType.ToString();
            chartTypeName =  chartTypeName.Remove(0, 2);
            if (Enum.IsDefined(typeof(ChartType), chartTypeName))
            {
                chartType = (PasteType)Enum.Parse(typeof(PasteType), chartTypeName, true);
            }
            return chartType;
        }

        public static XlPasteType ConvertBack(PasteType pasteType)
        {
            XlPasteType xlChartType = XlPasteType.xlPasteAll;
            string chartTypeName = pasteType.ToString();
            chartTypeName = chartTypeName.Insert(0, "xl");
            if (Enum.IsDefined(typeof(XlPasteType), chartTypeName))
            {
                xlChartType = (XlPasteType)Enum.Parse(typeof(XlPasteType), chartTypeName, true);
            }
            return xlChartType;
        }
    }
}
