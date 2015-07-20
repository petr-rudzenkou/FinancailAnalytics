using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Utils;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Screener.Base;
using FinancialAnalytics.Views.Screener.Events;
using FinancialAnalytics.Views.Screener.Interfaces;

namespace FinancialAnalytics.Views.Screener
{
    public class ScreenResultsViewModel : Screen, IScreenResultsViewModel, IHandle<ScreenCompletedEvent>, IHandle<ScreenerClosedEvent>
    {
        private readonly IScreenerResultsCollection _screenerResultsCollection;
        private readonly IExcelExporter _excelExporter;
        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private QuotesData _selectedScreenerQuote;
        private bool _isEnabled;
        private bool _isInPortfolio;

        private bool _asFormulas = false;
        private bool _asSelected = false;

        public ScreenResultsViewModel(IScreenerResultsCollection screenerResultsCollection, IEventAggregator eventAggregator, IExcelExporter excelExporter, IPortfolioQuotesCollection portfolioQuotesCollection, IViewsRenderer viewsRenderer)
        {
            _eventAggregator = eventAggregator;
            _eventAggregator.Subscribe(this);
            _screenerResultsCollection = screenerResultsCollection;
            _excelExporter = excelExporter;
            _portfolioQuotesCollection = portfolioQuotesCollection;
            _viewsRenderer = viewsRenderer;
            DisplayName = Resources.ViewsResources.ScreenResultsViewModel_DisplayName;
        }

        public BindableCollection<QuotesData> ScreenerQuotes
        {
            get { return _screenerResultsCollection.Quotes; }
        }

        public QuotesData SelectedScreenerQuote
        {
            get { return _selectedScreenerQuote; }
            set
            {
                _selectedScreenerQuote = value;
                NotifyOfPropertyChange(() => SelectedScreenerQuote);
            }
        }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                _isEnabled = value;
                NotifyOfPropertyChange(() => IsEnabled);
            }
        }

        public bool IsInPortfolio
        {
            get { return _isInPortfolio; }
            set
            {
                _isInPortfolio = value;
                NotifyOfPropertyChange(() => IsInPortfolio);
            }
        }

        public void AddToPortfolio()
        {
            if (SelectedScreenerQuote != null)
            {
                _portfolioQuotesCollection.Add(SelectedScreenerQuote);
            }
        }

        public void Insert()
        {
            List<QuotesData> quotes;
            if (_asSelected)
            {
                if (SelectedScreenerQuote == null)
                {
                    MessageBox.Show("Select a company");
                    return;
                }
                quotes = new List<QuotesData>() { SelectedScreenerQuote };
            }
            else
            {
                quotes = ScreenerQuotes.ToList();
            }

            if (!quotes.Any())
            {
                MessageBox.Show("Nothing to insert");
                return;
            }
                
            if (_asFormulas)
            {
                _excelExporter.InsertTickets(quotes.Select(x => x.Symbol));
            }
            else
            {
                _excelExporter.InsertData(quotes, QuotesProperties.DefaultQuoteProperties);
            }
        }

        public void Handle(ScreenCompletedEvent message)
        {
            if (message.HasResults)
            {
                IsEnabled = true;
            }
        }

        public void GetQuotes(string symbol)
        {
            _viewsRenderer.Show(ViewType.Quotes);
            _eventAggregator.Publish(new GetQuotesEvent() { Symbol = symbol });
        }

        public void GetChart(string symbol)
        {
            _viewsRenderer.Show(ViewType.Charts);
            _eventAggregator.Publish(new GetChartEvent() { Symbol = symbol });
        }

        public void ScreenResultsSelectionChanged(object sender, EventArgs e)
        {
            var quotesData = sender as QuotesData;
            if (quotesData != null)
            {
                IsInPortfolio = !PortfolioCacheProvider.PortfolioSymbols.Contains(quotesData.Symbol);
            }
            else
            {
                IsInPortfolio = false;
            }
        }

        public bool AsFormulas
        {
            get { return _asFormulas; }
            set
            {
                _asFormulas = value;
                NotifyOfPropertyChange(() => AsFormulas);
            }
        }

        public bool AsSelected
        {
            get { return _asSelected; }
            set
            {
                _asSelected = value;
                NotifyOfPropertyChange(() => AsSelected);
            }
        }

        public void RemoveQuotes(string symbol)
        {
            if (string.IsNullOrEmpty(symbol))
                return;

            _screenerResultsCollection.Remove(symbol);
            NotifyOfPropertyChange(() => ScreenerQuotes);
        }

        public void Handle(ScreenerClosedEvent message)
        {
            IsEnabled = false;
        }
    }
}
