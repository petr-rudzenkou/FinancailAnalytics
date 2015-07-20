
namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IFloor
    {
        IInterior Interior { get; }
        IChartFormat Format { get; }
        IBorder Border { get; }
    }
}
