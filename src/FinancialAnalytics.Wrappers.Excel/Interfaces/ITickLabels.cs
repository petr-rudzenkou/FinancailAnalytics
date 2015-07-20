using System;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ITickLabels
    {
        int Alignment { get; set; }
        string NumberFormat { get; set; }
        bool NumberFormatLinked { get; set; }
        int Offset { get; set; }
        IFont Font { get; }
		IChartFormat Format { get; }
    }
}
