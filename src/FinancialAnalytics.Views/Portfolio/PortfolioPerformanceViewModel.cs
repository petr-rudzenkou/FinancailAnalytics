using Caliburn.Micro;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;

namespace FinancialAnalytics.Views.Portfolio
{
    public class PortfolioPerformanceViewModel : PortfolioViewModelBase, IPortfolioPerformanceViewModel
    {
        public PortfolioPerformanceViewModel(IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
            : base(portfolioQuotesCollection, eventAggregator, viewsRenderer)
        {
            DisplayName = Resources.ViewsResources.Portfolio_PerfomanceView_Title;
        }        
    }
}
