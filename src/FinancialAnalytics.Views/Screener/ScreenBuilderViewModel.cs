using System;
using System.Collections.Generic;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.DataFacades.Screener;
using FinancialAnalytics.DataFacades.Screener.CriteriaGroups;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;
using FinancialAnalytics.Views.Screener.Base;
using FinancialAnalytics.Views.Screener.Events;
using FinancialAnalytics.Views.Screener.Interfaces;

namespace FinancialAnalytics.Views.Screener
{
    public class ScreenBuilderViewModel : Screen, IScreenBuilderViewModel, IHandle<CancelScreenEvent>, IHandle<ScreenerClosedEvent>
    {
        private readonly BindableCollection<CriteriaGroup> _criteriaGroups = new BindableCollection<CriteriaGroup>();
        private readonly StockScreenerDownload _stockScreenerDownload;
        private readonly IScreenerResultsCollection _screenerResultsCollection;
        private readonly IEventAggregator _eventAggregator;


        private bool _screeningEnabled = true;
        public ScreenBuilderViewModel(IScreenerResultsCollection screenerResultsCollection, IEventAggregator eventAggregator)
        {
            _screenerResultsCollection = screenerResultsCollection;
            _eventAggregator = eventAggregator;
            _eventAggregator.Subscribe(this);
            _stockScreenerDownload = new StockScreenerDownload();
            DisplayName = Resources.ViewsResources.ScreenBuilderViewModel_DisplayName;
            CreateCriteriaGroups();
        }

        public BindableCollection<CriteriaGroup> CriteriaGroups
        {
            get { return _criteriaGroups; }
        }

        private void CreateCriteriaGroups()
        {
            _criteriaGroups.Add(new CategoryGroup());
            _criteriaGroups.Add(new ShareDataGroup());
            _criteriaGroups.Add(new SalesAndProfitabilityGroup());
            _criteriaGroups.Add(new ValuationRatiosGroup());
        }

        public bool ScreeningEnabled
        {
            get { return _screeningEnabled; }
            set
            {
                _screeningEnabled = value;
                NotifyOfPropertyChange(() => ScreeningEnabled);
            }
        }

        public void RunScreen()
        {
            //_criteriaGroups.Refresh();
            ScreeningEnabled = false;
            _screenerResultsCollection.Clear();

            var criterias = new List<Criteria>();
            foreach (var group in _criteriaGroups)
            {
                criterias.AddRange(group.CriteriaFilters);
            }

            _stockScreenerDownload.AsyncDownloadCompleted += StockScreenerDownloadCompleted;
            _stockScreenerDownload.DownloadAsync(criterias, null);
            _eventAggregator.Publish(new RunScreenEvent());
        }

        private void StockScreenerDownloadCompleted(DownloadClient<StockScreenerResult> sender, DownloadCompletedEventArgs<StockScreenerResult> e)
        {
            var result = new StockScreenerResult(new QuotesData[0]);
            try
            {
                result = e.Response.Result;
                if (result != null)
                {
                    if (result.Items.Length > 0)
                    {
                        _screenerResultsCollection.Add(result.Items);
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                result = new StockScreenerResult(new QuotesData[0]);
                result.FinalResponse = true;
                MessageBox.Show("Error");
            }
            finally
            {
                if (result.FinalResponse)
                {
                    ScreeningEnabled = true;
                    var screenCompletedEvent = new ScreenCompletedEvent();
                    if (result.Items.Length > 0)
                    {
                        screenCompletedEvent.HasResults = true;
                    }
                    _stockScreenerDownload.AsyncDownloadCompleted -= StockScreenerDownloadCompleted;
                    _eventAggregator.Publish(screenCompletedEvent);
                }
            }
        }

        public void ClearFilters()
        {
            _criteriaGroups.Clear();
            CreateCriteriaGroups();
        }

        public void Handle(CancelScreenEvent message)
        {
            _stockScreenerDownload.AsyncDownloadCompleted -= StockScreenerDownloadCompleted;
            _stockScreenerDownload.CancelAsyncAll();
            ScreeningEnabled = true;
        }

        public void Handle(ScreenerClosedEvent message)
        {
            ClearFilters();
        }
    }
}
