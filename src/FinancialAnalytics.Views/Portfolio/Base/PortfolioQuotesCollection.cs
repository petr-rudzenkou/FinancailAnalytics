using System;
using System.Collections.Generic;
using System.Linq;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Utils;
using FinancialAnalytics.Views.Base;

namespace FinancialAnalytics.Views.Portfolio.Base
{
    public class PortfolioQuotesCollection : QuotesCollectionBase, IPortfolioQuotesCollection
    {
        private readonly QuotesDownload _quotesDownload;
        public event EventHandler InitializeStarted;
        public event EventHandler InitializeCompleted;

        public PortfolioQuotesCollection()
        {
            _quotesDownload = new QuotesDownload();
            _quotesDownload.AsyncDownloadCompleted += DownloadCompleted;
        }

        public void Initialize()
        {
            try
            {
                var quotesToDownload = new List<string>();
                foreach (var symbol in PortfolioCacheProvider.PortfolioSymbols)
                {
                    if (!Quotes.Any(x => x.Symbol.Equals(symbol)))
                    {
                        quotesToDownload.Add(symbol);
                    }
                }
                if (quotesToDownload.Count > 0)
                {
                    _quotesDownload.DownloadAsync(quotesToDownload, null);
                    OnInitializeStarted();
                }
            }
            catch (Exception ex)
            { }
        }

        public override void Add(QuotesData quotesData)
        {
            base.Add(quotesData);
            PortfolioCacheProvider.Add(quotesData.Symbol);
        }

        public override void Remove(string symbol)
        {
            base.Remove(symbol);
            PortfolioCacheProvider.Remove(symbol);
        }

        private void DownloadCompleted(DownloadClient<QuotesResult> sender, DownloadCompletedEventArgs<QuotesResult> e)
        {
            try
            {
                var response = e.Response;
                if (response != null)
                {
                    var items = response.Result.Items;
                    if (items.Any())
                    {
                        foreach (var item in items)
                        {
                            Quotes.Add(item);
                            if (!PortfolioCacheProvider.PortfolioSymbols.Any(x => x.Equals(item.Symbol)))
                            {
                                PortfolioCacheProvider.Add(item.Symbol);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                OnInitializeCompleted();
            }
        }

        private void OnInitializeCompleted()
        {
            var temp = InitializeCompleted;
            if (temp != null)
            {
                temp(null, null);
            }
        }

        private void OnInitializeStarted()
        {
            var temp = InitializeStarted;
            if (temp != null)
            {
                temp(null, null);
            }
        }


        public void Add(string symbol)
        {
            try
            {
                if (!Quotes.Any(x => x.Symbol.Equals(symbol)))
                {
                    _quotesDownload.DownloadAsync(new[] { symbol }, null);
                    OnInitializeStarted();
                }
            }
            catch (Exception ex)
            { }
        }
    }
}
