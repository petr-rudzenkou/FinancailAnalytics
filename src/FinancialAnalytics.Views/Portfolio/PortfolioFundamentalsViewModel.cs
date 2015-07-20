using Caliburn.Micro;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;

namespace FinancialAnalytics.Views.Portfolio
{
    public class PortfolioFundamentalsViewModel : PortfolioViewModelBase, IPortfolioFundamentalsViewModel
    {
        public PortfolioFundamentalsViewModel(IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
            : base(portfolioQuotesCollection, eventAggregator, viewsRenderer)
        {
            DisplayName = Resources.ViewsResources.Portfolio_FundamentalsView_Title;
        }
    }
}
