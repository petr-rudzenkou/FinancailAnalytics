using System.Collections.Generic;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    public interface IEntitiesCollectionWrapper<TCollection, TItem> : IEntityWrapper<TCollection>, IEnumerable<TItem>
    {
        TItem this[int index] { get; }

        int Count { get; }

        void FullDispose();
    }
}
