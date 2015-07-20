using System;
using FinancialAnalytics.Views.Base;

namespace FinancialAnalytics.Views.Portfolio.Base
{
    public interface IPortfolioQuotesCollection : IQuotesCollectionBase
    {
        void Initialize();
        event EventHandler InitializeStarted;
        event EventHandler InitializeCompleted;
        void Add(string symbol);
    }
}
