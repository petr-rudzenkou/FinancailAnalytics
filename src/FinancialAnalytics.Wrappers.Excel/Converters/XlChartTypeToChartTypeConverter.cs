using System;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlChartTypeToChartTypeConverter
    {
        public static ChartType Convert(XlChartType xlChartType)
        {
            //fix IBLINK-455 Fill options of charts are changed if to apply chart attributes copied from line, scatter, radar charts in which line options are changed
            //This code return when chart type is line
            if (xlChartType.ToString() == "-4111")
            {
                xlChartType = XlChartType.xlLine;
            }
            ChartType chartType = ChartType.Chart_ColumnClustered;
            string chartTypeName = xlChartType.ToString();
            chartTypeName =  chartTypeName.Remove(0, 2);
            chartTypeName = chartTypeName.Insert(0, "Chart_");
            if (Enum.IsDefined(typeof(ChartType), chartTypeName))
            {
                chartType = (ChartType)Enum.Parse(typeof(ChartType), chartTypeName, true);
            }
            return chartType;
        }

        public static XlChartType ConvertBack(ChartType chartType)
        {
            XlChartType xlChartType = XlChartType.xlColumnClustered;
            string chartTypeName = chartType.ToString();
            chartTypeName = chartTypeName.Remove(0, 6);
            chartTypeName = chartTypeName.Insert(0, "xl");
            if (Enum.IsDefined(typeof(XlChartType), chartTypeName))
            {
                xlChartType = (XlChartType)Enum.Parse(typeof(XlChartType), chartTypeName, true);
            }
            return xlChartType;
        }
    }
}
