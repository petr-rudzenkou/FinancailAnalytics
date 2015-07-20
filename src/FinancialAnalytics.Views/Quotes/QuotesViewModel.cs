using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.ProgressBar;
using FinancialAnalytics.Views.Quotes.Interfaces;
using Action = System.Action;

namespace FinancialAnalytics.Views.Quotes
{
    public class QuotesViewModel : Conductor<IScreen>.Collection.OneActive, IQuotesViewModel, IHandle<GetQuotesEvent>
    {
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private readonly IExcelExporter _excelExporter;
        private readonly IProgressBarService _progressBarService;
        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;
        private readonly QuotesDownload _quotesDownload;
        private IScreen _selectedView;
        private FrameworkElement _view; //TODO:remove logic of handlig text boxes into UI (triggers)
        private bool _asFormulas = false;
        private bool _asSelected = false;

        public QuotesViewModel(IExcelExporter excelExporter, IProgressBarService progressBarService,
            IPortfolioQuotesCollection portfolioQuotesCollection, IEventAggregator eventAggregator,
            IViewsRenderer viewsRenderer)
        {
            _eventAggregator = eventAggregator;
            _eventAggregator.Subscribe(this);
            _viewsRenderer = viewsRenderer;
            _excelExporter = excelExporter;
            _progressBarService = progressBarService;
            _progressBarService.Cancelled += QuotesDownloadCanceled;
            _portfolioQuotesCollection = portfolioQuotesCollection;
            _quotesDownload = new QuotesDownload();
            _quotesDownload.AsyncDownloadCompleted += DownloadCompleted;
            DisplayName = Resources.ViewsResources.Qoutes_Window_Title;
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

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            _view = view as FrameworkElement;
        }

        public void AddQuotes(string symbols)
        {
            try
            {
                var trimedSymbols = symbols.Trim();
                if (string.IsNullOrEmpty(trimedSymbols))
                    return;

                var arraySymbols = trimedSymbols.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                _quotesDownload.DownloadAsync(arraySymbols, null);
                ClearSymbolsTextBox();
                _progressBarService.Show(this);
            }
            catch (Exception ex)
            {
                //_log.Error(ex);
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

        public void Insert()
        {
            var quotesInfoModels = new List<QuotesInfoViewModel>();
            foreach (var item in Items)
            {
                var model = item as QuotesInfoViewModel;
                if (model != null)
                {
                    quotesInfoModels.Add(model);
                }
            }
            if (quotesInfoModels.Count > 0)
            {
                List<QuotesData> quotes;
                if (_asSelected)
                {
                    var quote = quotesInfoModels.FirstOrDefault(x => x.IsActive);
                    if (quote != null)
                    {
                        quotes = new List<QuotesData>() { quote.QuotesData };
                    }
                    else
                    {
                        quotes = new List<QuotesData>();
                    }
                }
                else
                {
                    quotes = quotesInfoModels.Select(x => x.QuotesData).ToList();
                }

                if (!quotes.Any())
                {
                    MessageBox.Show("Nothing to insert");
                    return;
                }

                if (_asFormulas)
                {
                    var tickets = quotes.Select(x => x.Symbol);
                    _excelExporter.InsertTickets(tickets);
                }
                else
                {
                    _excelExporter.InsertData(quotes, QuotesProperties.DefaultQuoteProperties);
                }
            }
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
                        foreach (var quote in items)
                        {
                            var item = GetModel(quote.Symbol);
                            if (item != null)
                            {
                                item.QuotesData = quote;
                            }
                            else
                            {
                                var quotesInfoViewModel = new QuotesInfoViewModel(_portfolioQuotesCollection, _eventAggregator, _viewsRenderer);
                                quotesInfoViewModel.QuotesData = quote;
                                Items.Add(quotesInfoViewModel);
                            }
                        }
                        NotifyOfPropertyChange(() => Items);
                        SelectedView = Items.LastOrDefault();
                    }
                    else
                    {
                        MessageBox.Show("There is no such ticket.");
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                _progressBarService.Close();
            }
        }

        public void CloseItem(string displayName) //Consider of binding to symbol
        {
            var itemToClose = GetModel(displayName);
            if (itemToClose != null)
            {
                int index = Items.IndexOf(itemToClose);
                Items.RemoveAt(index);
                NotifyOfPropertyChange(() => Items);
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

        private QuotesInfoViewModel GetModel(string symbol)
        {
            var quotesInfoModels = new List<QuotesInfoViewModel>();
            foreach (var item in Items)
            {
                var model = item as QuotesInfoViewModel;
                if (model != null)
                {
                    quotesInfoModels.Add(model);
                }
            }
            return quotesInfoModels.FirstOrDefault(x => x.QuotesData.Symbol == symbol);
        }

        private void QuotesDownloadCanceled(object sender, EventArgs e)
        {
            _quotesDownload.CancelAsyncAll();
        }

        public void Handle(GetQuotesEvent message)
        {
            var symbol = message.Symbol;
            var model = GetModel(symbol);
            if (model == null)
            {
                _quotesDownload.DownloadAsync(new[] { symbol }, null);
                _progressBarService.Show(this);
            }
            else
            {
                SelectedView = model;
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
    }
}
