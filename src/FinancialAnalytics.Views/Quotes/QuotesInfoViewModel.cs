using System.Linq;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Utils;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Quotes.Interfaces;

namespace FinancialAnalytics.Views.Quotes
{
    public class QuotesInfoViewModel : Screen, IQuotesInfoViewModel
    {
        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private QuotesData _quotesData;

        public QuotesInfoViewModel(IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
        {
            _portfolioQuotesCollection = portfolioQuotesCollection;
            _eventAggregator = eventAggregator;
            _viewsRenderer = viewsRenderer;
        }
        public QuotesData QuotesData
        {
            get { return _quotesData; }
            set
            {
                _quotesData = value;
                DisplayName = _quotesData.Symbol;
                NotifyOfPropertyChange(() => QuotesData);
            }
        }

        public void AddToPortfolio()
        {
            _portfolioQuotesCollection.Add(QuotesData);
            Refresh();
        }

        public void GetChart(string symbol)
        {
            _viewsRenderer.Show(ViewType.Charts);
            _eventAggregator.Publish(new GetChartEvent() { Symbol = symbol });
        }
    }
}
