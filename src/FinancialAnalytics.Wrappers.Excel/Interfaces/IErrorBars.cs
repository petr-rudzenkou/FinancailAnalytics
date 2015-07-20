using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IErrorBars : IEntityWrapper<IErrorBars>
    {
        IBorder Border { get; }
    }
}
