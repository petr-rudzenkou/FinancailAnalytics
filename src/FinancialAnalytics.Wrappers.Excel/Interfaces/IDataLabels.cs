
namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IDataLabels
    {
        IFont Font { get; }
        IChartFormat Format { get; }
        IBorder Border { get; }
    }
}
