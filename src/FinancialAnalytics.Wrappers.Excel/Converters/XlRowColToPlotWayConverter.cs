using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlRowColToPlotWayConverter
    {
        public static RowCol Convert(XlRowCol xlRowCol)
        {
            RowCol rowCol;
            switch (xlRowCol)
            {
                case XlRowCol.xlRows:
                    rowCol = RowCol.Rows;
                    break;
                default:
                    rowCol = RowCol.Columns;
                    break;
            }
            return rowCol;
        }

        public static XlRowCol ConvertBack(RowCol rowCol)
        {
            XlRowCol xlRowCol;
            switch (rowCol)
            {
                case RowCol.Rows:
                    xlRowCol = XlRowCol.xlRows;
                    break;
                default:
                    xlRowCol = XlRowCol.xlColumns;
                    break;
            }
            return xlRowCol;
        }
    }
}
