using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISeriesLines : IEntityWrapper<ISeriesLines>
    {
        IBorder Border { get; }
    }
}
