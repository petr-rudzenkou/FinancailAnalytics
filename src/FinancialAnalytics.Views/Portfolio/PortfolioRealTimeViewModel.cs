using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;

namespace FinancialAnalytics.Views.Portfolio
{
    /// <summary>
    /// Currently not used.
    /// </summary>
    public class PortfolioRealTimeViewModel : PortfolioViewModelBase, IPortfolioRealTimeViewModel
    {
        public PortfolioRealTimeViewModel(IPortfolioQuotesCollection quotesCollection)
            :base(quotesCollection)
        {
            DisplayName = Resources.ViewsResources.Portfolio_RealTimeView_Title;
        }
    }
}
