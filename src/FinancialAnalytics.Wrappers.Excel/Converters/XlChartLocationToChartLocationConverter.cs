using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{

    public class XlChartLocationToChartLocationConverter
    {
        public static ChartLocation Convert(XlChartLocation xlChartLocation)
        {
            ChartLocation chartLocation;
            switch (xlChartLocation)
            {
                case XlChartLocation.xlLocationAsNewSheet:
                    chartLocation = ChartLocation.LocationAsNewSheet;
                    break;
                case XlChartLocation.xlLocationAsObject:
                    chartLocation = ChartLocation.LocationAsObject;
                    break;
                default:
                    chartLocation = ChartLocation.LocationAutomatic;
                    break;
            }
            return chartLocation;
        }

        public static XlChartLocation ConvertBack(ChartLocation chartLocation)
        {
            XlChartLocation xlChartLocation;
            switch (chartLocation)
            {
                case ChartLocation.LocationAsNewSheet:
                    xlChartLocation = XlChartLocation.xlLocationAsNewSheet;
                    break;
                case ChartLocation.LocationAsObject:
                    xlChartLocation = XlChartLocation.xlLocationAsObject;
                    break;
                default:
                    xlChartLocation = XlChartLocation.xlLocationAutomatic;
                    break;
            }
            return xlChartLocation;
        }
    }
}
