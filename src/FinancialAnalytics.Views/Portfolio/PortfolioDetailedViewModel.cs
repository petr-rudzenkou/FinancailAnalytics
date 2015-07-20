using Caliburn.Micro;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;

namespace FinancialAnalytics.Views.Portfolio
{
    public class PortfolioDetailedViewModel : PortfolioViewModelBase, IPortfolioDetailedViewModel
    {
        public PortfolioDetailedViewModel(IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
            : base(portfolioQuotesCollection, eventAggregator, viewsRenderer)
        {
            DisplayName = Resources.ViewsResources.Portfolio_DetailedView_Title;
        }
    }
}
