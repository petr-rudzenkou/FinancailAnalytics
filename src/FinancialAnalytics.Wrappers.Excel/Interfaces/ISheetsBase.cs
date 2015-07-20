using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISheetsBase<TCollection, TItem> : IEntitiesCollectionWrapper<TCollection, TItem>
    {
        TItem AddLast();
        TItem AddFirst();
    }
}
