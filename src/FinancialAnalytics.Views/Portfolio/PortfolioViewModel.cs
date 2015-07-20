using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;
using System.Windows.Controls;
using FinancialAnalytics.Views.ProgressBar;
using Action = System.Action;

namespace FinancialAnalytics.Views.Portfolio
{
    public class PortfolioViewModel : Conductor<IScreen>.Collection.OneActive, IPortfolioViewModel
    {
        private readonly IPortfolioBasicViewModel _portfolioBasicViewModel;
        private readonly IPortfolioDetailedViewModel _portfolioDetailedViewModel;
        private readonly IPortfolioFundamentalsViewModel _portfolioFundamentalsViewModel;
        private readonly IPortfolioPerformanceViewModel _portfolioPerformanceViewModel;

        private IScreen _selectedView;

        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;

        private readonly IEventAggregator _eventAggregator;
        private readonly IExcelExporter _excelExporter;
        private readonly IProgressBarService _progressBarService;

        private readonly QuotesDownload _quotesDownload;
        private FrameworkElement _view; //TODO:remove logic of handlig text boxes into UI (triggers)

        private bool _asFormulas = false;
        private bool _asSelected = false;

        public PortfolioViewModel(
            IPortfolioBasicViewModel portfolioBasicViewModel,
            IPortfolioDetailedViewModel portfolioDetailedViewModel,
            IPortfolioPerformanceViewModel portfolioPerformanceViewModel,
            IPortfolioFundamentalsViewModel portfolioFundamentalsViewModel,
            IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator,
            IExcelExporter excelExporter,
            IProgressBarService progressBarService)
        {
            _portfolioBasicViewModel = portfolioBasicViewModel;
            _portfolioDetailedViewModel = portfolioDetailedViewModel;
            _portfolioPerformanceViewModel = portfolioPerformanceViewModel;
            _portfolioFundamentalsViewModel = portfolioFundamentalsViewModel;
            _portfolioQuotesCollection = portfolioQuotesCollection;
            _portfolioQuotesCollection.InitializeCompleted += OnInitializePortfolioCompleted;
            _portfolioQuotesCollection.InitializeStarted += OnInitializePortfolioStarted;
            _eventAggregator = eventAggregator;
            _eventAggregator.Subscribe(this);

            _excelExporter = excelExporter;
            _progressBarService = progressBarService;
            _progressBarService.Cancelled += QuotesDownloadCanceled;

            _quotesDownload = new QuotesDownload();
            _quotesDownload.AsyncDownloadCompleted += DownloadCompleted;

            DisplayName = Resources.ViewsResources.Portfolio_WindowTitle;
        }

        public IScreen SelectedView
        {
            get { return _selectedView; }
            set
            {
                if (_selectedView == value)
                {
                    return;
                }
                ChangeActiveItem(value, false);
            }
        }

        protected override void ChangeActiveItem(IScreen newItem, bool closePrevious)
        {
            UpdateLayout(newItem);
            base.ChangeActiveItem(newItem, closePrevious);
        }

        private void UpdateLayout(IScreen newItem)
        {
            _selectedView = newItem;
            NotifyOfPropertyChange(() => SelectedView);
        }

        private void SetDataSource()
        {
            Items.Clear();
            Items.Add(_portfolioBasicViewModel);
            Items.Add(_portfolioFundamentalsViewModel);
            Items.Add(_portfolioPerformanceViewModel);
            Items.Add(_portfolioDetailedViewModel);

            SelectedView = Items.FirstOrDefault();
        }

        public int QuotesCount
        {
            get { return _portfolioQuotesCollection.Quotes.Count; }
        }

        public void AddQuotes(string symbols)
        {
            try
            {
                ClearSymbolsTextBox();

                var trimedSymbols = symbols.Trim();
                if (string.IsNullOrEmpty(trimedSymbols))
                    return;

                var arraySymbols = trimedSymbols.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                _quotesDownload.DownloadAsync(arraySymbols, null);
                _progressBarService.Show(this);
            }
            catch (Exception ex)
            { }
        }

        public void Insert()
        {
            List<QuotesData> quotes;
            if (_asSelected)
            {
                var activeItem = SelectedView as PortfolioViewModelBase;
                if (activeItem == null)
                {
                    MessageBox.Show("Select a quote");
                    return;
                }
                if (activeItem.SelectedQuote == null)
                {
                    MessageBox.Show("Select a quote");
                    return;
                }
                quotes = new List<QuotesData>() { activeItem.SelectedQuote };
            }
            else
            {
                quotes = _portfolioQuotesCollection.Quotes.ToList();
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

        public void ExecuteAddQuotes(ActionExecutionContext context)
        {
            var keyArgs = context.EventArgs as KeyEventArgs;
            if (keyArgs == null)
                return;

            if (keyArgs.Key != Key.Enter)
                return;

            var textBox = context.Source as TextBox;
            if (textBox != null)
            {
                AddQuotes(textBox.Text);
            }
        }

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            SetDataSource();
            _view = view as FrameworkElement;
        }

        protected override void OnActivate()
        {
            base.OnActivate();
            _portfolioQuotesCollection.Initialize();
        }

        private void DownloadCompleted(object sender, DownloadCompletedEventArgs<QuotesResult> args)
        {
            try
            {
                var response = args.Response;
                if (response != null)
                {
                    var items = response.Result.Items;
                    if (items.Any())
                    {
                        foreach (var item in items)
                        {
                            _portfolioQuotesCollection.Add(item);
                        }
                        NotifyOfPropertyChange(() => QuotesCount);
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                _progressBarService.Close();
            }
        }

        private void ClearSymbolsTextBox()
        {
            if (_view != null)
            {
                _view.Dispatcher.BeginInvoke(new Action(() =>
                {
                    var textBox = _view.FindName("Symbols") as TextBox;
                    if (textBox != null)
                    {
                        textBox.Clear();
                    }
                }));
            }
        }

        private void QuotesDownloadCanceled(object sender, EventArgs e)
        {
            _quotesDownload.CancelAsyncAll();
        }

        private void OnInitializePortfolioStarted(object sender, EventArgs e)
        {
            if (!IsActive)
                return;

            _progressBarService.Show(this);
        }

        private void OnInitializePortfolioCompleted(object sender, EventArgs e)
        {
            _progressBarService.Close();
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
    }
}
