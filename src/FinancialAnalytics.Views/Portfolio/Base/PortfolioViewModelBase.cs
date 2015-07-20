using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Views.Events;

namespace FinancialAnalytics.Views.Portfolio.Base
{
    public class PortfolioViewModelBase : Screen
    {
        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private QuotesData _selectedQuote;

        public PortfolioViewModelBase(IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
        {
            _portfolioQuotesCollection = portfolioQuotesCollection;
            _eventAggregator = eventAggregator;
            _viewsRenderer = viewsRenderer;
        }

        public BindableCollection<QuotesData> Quotes
        {
            get { return _portfolioQuotesCollection.Quotes; }
        }

        public QuotesData SelectedQuote
        {
            get { return _selectedQuote; }
            set
            {
                _selectedQuote = value;
                NotifyOfPropertyChange(() => SelectedQuote);
            }
        }

        public void RemoveQuotes(string symbol)
        {
            if (string.IsNullOrEmpty(symbol))
                return;

            _portfolioQuotesCollection.Remove(symbol);
            NotifyOfPropertyChange(() => Quotes);
        }

        public void GetChart(string symbol)
        {
            _viewsRenderer.Show(ViewType.Charts);
            _eventAggregator.Publish(new GetChartEvent() { Symbol = symbol });
        }

        public void GetQuotes(string symbol)
        {
            _viewsRenderer.Show(ViewType.Quotes);
            _eventAggregator.Publish(new GetQuotesEvent() { Symbol = symbol });
        }
    }
}
