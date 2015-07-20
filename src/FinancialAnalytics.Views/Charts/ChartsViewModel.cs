using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades;
using FinancialAnalytics.DataFacades.Charts;
using FinancialAnalytics.Utils;
using FinancialAnalytics.Views.Charts.Interfaces;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.Portfolio.Base;
using Action = System.Action;

namespace FinancialAnalytics.Views.Charts
{
    public class ChartsViewModel : Conductor<IScreen>.Collection.OneActive, IChartsViewModel, IHandle<GetChartEvent>
    {
        private readonly IExcelExporter _excelExporter;
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private readonly IPortfolioQuotesCollection _portfolioQuotesCollection;
        private IScreen _selectedView;
        private FrameworkElement _view; //TODO:remove logic of handlig text boxes into UI (triggers)

        public ChartsViewModel(IExcelExporter excelExporter, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer, IPortfolioQuotesCollection portfolioQuotesCollection)
        {
            eventAggregator.Subscribe(this);
            _eventAggregator = eventAggregator;
            _excelExporter = excelExporter;
            _viewsRenderer = viewsRenderer;
            _portfolioQuotesCollection = portfolioQuotesCollection;
            DisplayName = Resources.ViewsResources.Charts_Window_Title;
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
            base.ChangeActiveItem(newItem, closePrevious);
            UpdateLayout(newItem);
        }
        private void UpdateLayout(IScreen newItem)
        {
            _selectedView = newItem;
            NotifyOfPropertyChange(() => SelectedView);
            NotifyOfPropertyChange(() => ActiveSymbol);
            NotifyOfPropertyChange(() => IsAnyActive);
        }

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            _view = view as FrameworkElement;
        }

        public void CloseItem(string displayName) //Consider of binding to symbol
        {
            var itemToClose = GetModel(displayName);
            if (itemToClose != null)
            {
                int index = Items.IndexOf(itemToClose);
                Items.RemoveAt(index);
                NotifyOfPropertyChange(() => Items);
                NotifyOfPropertyChange(() => IsAnyActive);
                NotifyOfPropertyChange(() => ActiveSymbol);
            }
        }

        public void AddCharts(string symbols)
        {
            try
            {
                ClearSymbolsTextBox();

                var trimedSymbols = symbols.Trim();
                if (string.IsNullOrEmpty(trimedSymbols))
                    return;

                var arraySymbols = trimedSymbols.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                foreach (var chartId in arraySymbols)
                {
                    var model = CreateChartInfoViewModel(chartId);
                    Items.Add(model);
                }
                NotifyOfPropertyChange(() => Items);
                SelectedView = Items.LastOrDefault();
            }
            catch (Exception ex)
            { }
        }

        public void ExecuteAddCharts(ActionExecutionContext context)
        {
            var keyArgs = context.EventArgs as KeyEventArgs;
            if (keyArgs == null)
                return;

            if (keyArgs.Key != Key.Enter)
                return;

            var textBox = context.Source as TextBox;
            if (textBox != null)
            {
                AddCharts(textBox.Text);
            }
        }

        public void Insert()
        {
            var activeChart = Items.FirstOrDefault(x => x.IsActive);
            if (activeChart != null)
            {
                var chartsInfoViewModel = activeChart as ChartsInfoViewModel;
                if (chartsInfoViewModel != null)
                {
                    _excelExporter.InsertImage(chartsInfoViewModel.Chart);
                }
            }
        }

        private ChartsInfoViewModel GetModel(string symbol)
        {
            var quotesInfoModels = new List<ChartsInfoViewModel>();
            foreach (var item in Items)
            {
                var model = item as ChartsInfoViewModel;
                if (model != null)
                {
                    quotesInfoModels.Add(model);
                }
            }
            return quotesInfoModels.FirstOrDefault(x => x.Symbol == symbol);
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

        private ChartsInfoViewModel CreateChartInfoViewModel(string symbol)
        {
            var settings = new ChartDownloadSettings()
            {
                ID = symbol,
                TimeSpan = ChartTimeSpan.c1Y,
                SimplifiedImage = false,
                ImageSize = ChartImageSize.Large,
            };

            var chartsInfoViewModel = new ChartsInfoViewModel(symbol)
            {
                Chart = new BitmapImage(new Uri(settings.GetUrl()))
            };

            return chartsInfoViewModel;
        }

        public void Handle(GetChartEvent message)
        {
            var symbol = message.Symbol;
            var model = GetModel(symbol);
            if (model == null)
            {
                model = CreateChartInfoViewModel(symbol);
                Items.Add(model);
                SelectedView = model;
            }
            else
            {
                SelectedView = model;
            }
        }

        public void GetQuotes()
        {
            var activeModel = SelectedView as ChartsInfoViewModel;
            if (activeModel != null)
            {
                _viewsRenderer.Show(ViewType.Quotes);
                _eventAggregator.Publish(new GetQuotesEvent() { Symbol = activeModel.Symbol });
            }
        }

        public void AddToPortfolio()
        {
            var activeModel = SelectedView as ChartsInfoViewModel;
            if (activeModel != null)
            {
                _portfolioQuotesCollection.Add(activeModel.Symbol);
                NotifyOfPropertyChange(() => SelectedView);
            }
        }

        public Visibility IsAnyActive
        {
            get
            {
                if (Items.Any())
                    return Visibility.Visible;
                return Visibility.Hidden;
            }
        }

        public string ActiveSymbol
        {
            get
            {
                string symbol = string.Empty;
                var activeModel = SelectedView as ChartsInfoViewModel;
                if (activeModel != null)
                {
                    symbol = activeModel.Symbol;
                }
                return symbol;
            }
        }
    }
}
