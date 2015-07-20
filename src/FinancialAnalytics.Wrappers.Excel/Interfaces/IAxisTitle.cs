using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IAxisTitle : IEntityWrapper<IAxisTitle>
    {
        string Text { get; set; }

        double Top { get; set; }

        double Left { get; set; }

        object Orientation { get; set; }

        IFont Font { get; }

		IChartFormat Format { get; }

        IBorder Border { get; }
    }
}
