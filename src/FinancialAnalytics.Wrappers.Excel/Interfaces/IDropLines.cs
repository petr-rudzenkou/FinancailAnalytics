using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IDropLines : IEntityWrapper<IDropLines>
    {
        string Name { get; }
        IBorder Border { get; }
    }
}
