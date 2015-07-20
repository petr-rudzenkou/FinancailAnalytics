using System;
using System.Linq;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.XChangeRates;
using FinancialAnalytics.DataFacades.XChangeRates.Metadata;
using FinancialAnalytics.Views.ProgressBar;
using FinancialAnalytics.Views.XChangeRates.Interfaces;

namespace FinancialAnalytics.Views.XChangeRates
{
    public class XChangeRatesViewModel : Screen, IXChangeRatesViewModel
    {
        private const string CHART_FORMAT_URL = "http://chart.finance.yahoo.com/instrument/1.0/{0}=X/chart;range=5d/image;size=300x170?region=US&lang=en-US&scheme=gsbeta";
        private readonly string[] _pairs = new[]
        {
            "USDBYR",
            "EURBYR",
            "RUBBYR",
            "EURUSD",
            "UAHBYR",
            "GBPBYR",
            "CHFBYR",
            "JPYBYR"
        };

        private readonly XChangeRatesDownload _xChangeRatesDownload;
        private readonly IProgressBarService _progressBarService;
        private readonly BindableCollection<XChangeRate> _rates = new BindableCollection<XChangeRate>();
        private XChangeRate _selectedRate;
        private Uri _xChangeRatesChart;

        public XChangeRatesViewModel(IProgressBarService progressBarService)
        {
            _progressBarService = progressBarService;
            DisplayName = Resources.ViewsResources.XchangeRates_Window_Title;
            _xChangeRatesDownload = new XChangeRatesDownload();
            _xChangeRatesDownload.AsyncDownloadCompleted += XChangeRatesAsyncDownloadCompleted;
        }

        public BindableCollection<XChangeRate> Rates
        {
            get { return _rates; }
        }

        private void Load()
        {
            _xChangeRatesDownload.DownloadAsync(_pairs);
            _progressBarService.Show(this);
        }

        public XChangeRate SelectedRate
        {
            get { return _selectedRate; }
            set
            {
                _selectedRate = value;
                NotifyOfPropertyChange(() => SelectedRate);
            }
        }

        public Uri XChangeRatesChart
        {
            get { return _xChangeRatesChart; }
            set
            {
                _xChangeRatesChart = value;
                NotifyOfPropertyChange(() => XChangeRatesChart);
            }
        }

        protected override void OnViewAttached(object view, object context)
        {
            base.OnViewAttached(view, context);
            
        }

        protected override void OnActivate()
        {
            base.OnActivate();
            Load();
        }

        public void RatesSelectionChanged(object sender, EventArgs e)
        {
            var xChangeRate = sender as XChangeRate;
            if(xChangeRate != null)
            {
                SetupChart(xChangeRate.Id);
            }
        }

        private void SetupChart(string id)
        {
            XChangeRatesChart = new Uri(string.Format(CHART_FORMAT_URL, !string.IsNullOrEmpty(id) ? id : string.Empty));
        }

        private void XChangeRatesAsyncDownloadCompleted(DownloadClient<XChangeRatesResult> sender, DownloadCompletedEventArgs<XChangeRatesResult> e)
        {
            try
            {
                var response = e.Response;
                if (response != null)
                {
                    var items = response.Result.Items;
                    if (items.Any())
                    {
                        Rates.Clear();
                        Rates.AddRange(items);
                        if (SelectedRate == null)
                        {
                            SelectedRate = Rates.FirstOrDefault();
                            if (SelectedRate != null)
                            {
                                SetupChart(SelectedRate.Id);
                            }
                        }
                        NotifyOfPropertyChange(() => Rates);
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
    }
}
