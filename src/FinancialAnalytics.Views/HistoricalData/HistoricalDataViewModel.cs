using System;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.HistoricalData;
using FinancialAnalytics.Views.Events;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.HistoricalData.Interfaces;
using FinancialAnalytics.Views.ProgressBar;

namespace FinancialAnalytics.Views.HistoricalData
{
    public class HistoricalDataViewModel : Screen, IHistoricalDataViewModel
    {
        private readonly IEventAggregator _eventAggregator;
        private readonly IViewsRenderer _viewsRenderer;
        private readonly IExcelExporter _excelExporter;
        private readonly IProgressBarService _progressBarService;
        private readonly BindableCollection<DataFacades.HistoricalData.HistoricalData> _historicalDatas = new BindableCollection<DataFacades.HistoricalData.HistoricalData>();
        private readonly HistoricalDataDownload _historicalDataDownload;

        public HistoricalDataViewModel(IExcelExporter excelExporter, IProgressBarService progressBarService, IEventAggregator eventAggregator, IViewsRenderer viewsRenderer)
        {
            _eventAggregator = eventAggregator;
            _viewsRenderer = viewsRenderer;
            _excelExporter = excelExporter;
            _progressBarService = progressBarService;
            _historicalDataDownload = new HistoricalDataDownload();
            _historicalDataDownload.AsyncDownloadCompleted += DownloadCompleted;
            DisplayName = Resources.ViewsResources.HistoricalData_Window_Title;
        }

        public BindableCollection<DataFacades.HistoricalData.HistoricalData> HistoricalDatas
        {
            get { return _historicalDatas; }
        }

        public void Insert()
        {
            try
            {
                _excelExporter.InsertData(HistoricalDatas);
            }
            catch(Exception ex)
            { }
        }

        public void GetPrices(string symbol, DateTime startDate, DateTime endDate)
        {
            try
            {
                if (string.IsNullOrEmpty(symbol))
                {
                    MessageBox.Show("Please, specify symbol");
                    return;
                }

                _historicalDataDownload.DownloadAsync(symbol, startDate, endDate);
                _progressBarService.Show(this);
            }
            catch(Exception ex)
            { }
        }

        private void DownloadCompleted(object sender, DownloadCompletedEventArgs<HistoricalDataResult> args)
        {
            try
            {
                var response = args.Response;
                if (response != null)
                {
                    var items = response.Result.Items;
                    if (items.Any())
                    {
                        HistoricalDatas.Clear();
                        foreach (var item in items)
                        {
                            HistoricalDatas.Add(item);
                        }
                        NotifyOfPropertyChange(() => HistoricalDatas);
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

        protected override void OnDeactivate(bool close)
        {
            base.OnDeactivate(close);
            if (close)
            {
                HistoricalDatas.Clear();
            }
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
