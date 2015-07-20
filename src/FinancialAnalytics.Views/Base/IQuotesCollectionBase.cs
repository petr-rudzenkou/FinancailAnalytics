using System.Collections.Generic;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Quotes;

namespace FinancialAnalytics.Views.Base
{
    public interface IQuotesCollectionBase
    {
        BindableCollection<QuotesData> Quotes { get; }
        void Add(QuotesData quotesData);
        void Add(IEnumerable<QuotesData> quotesData);
        void Remove(QuotesData quotesData);
        void Remove(IEnumerable<QuotesData> quotesData);
        void Remove(string symbol);
        void Remove(IEnumerable<string> symbols);
        void Clear();
    }
}
