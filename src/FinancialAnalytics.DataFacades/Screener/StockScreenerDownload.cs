using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.DataFacades.Screener.Criterias;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;
using FinancialAnalytics.DataFacades.Screener.Metedata;
using FinancialAnalytics.Utils;

namespace FinancialAnalytics.DataFacades.Screener
{
    public class StockScreenerDownload : DownloadClient<StockScreenerResult>
    {
        private readonly List<string> _industryIds = new List<string>();
        private readonly object _lock = new object();

        public Response<StockScreenerResult> Download(IEnumerable<Criteria> criterias, object userArgs)
        {
            var settings = new StockScreenerDownloadSettings(criterias.ToArray());
            return base.Download(settings);
        }
        public void DownloadAsync(IEnumerable<Criteria> criterias, object userArgs)
        {
            _industryIds.Clear();
            var criteria = criterias.FirstOrDefault(x => x.Name == "Industry");
            var industryCriteria = criteria as IndustryCriteria;
            if (industryCriteria != null && industryCriteria.SelectedIndustry != null)
            {
                if (industryCriteria.SelectedIndustry.Id.Equals("0")) //Any
                {
                    _industryIds.AddRange(Industries.All.Select(x => x.Id).ToArray());
                }
                else
                {
                    var industryId = industryCriteria.SelectedIndustry.Id;
                    _industryIds.Add(industryId);
                }
            }
            else
            {
                _industryIds.AddRange(Industries.All.Select(x => x.Id).ToArray());
            }
            foreach (var id in _industryIds)
            {
                var settings = new StockScreenerDownloadSettings(criterias.ToArray());
                settings.IndustryId = id;
                base.DownloadAsync(settings, null);
            }

        }

        protected override StockScreenerResult ConvertResult(ConnectionInfo connInfo, System.IO.Stream stream, SettingsBase settings)
        {
            var quotes = new List<QuotesData>();
            try
            {
                StockScreenerResult screenerResult;
                string result = DataFacadesHelper.StreamToString(stream);
                IEnumerable<QuotesData> quotesData = ImportExport.XmlToQuoteData(result);
                lock (_lock)
                {
                    var stockScreenerSettings = settings as StockScreenerDownloadSettings;
                    if (stockScreenerSettings != null)
                    {
                        _industryIds.Remove(stockScreenerSettings.IndustryId);
                        quotes.AddRange(FilterQuotes(quotesData, stockScreenerSettings.Criterias));
                    }
                    screenerResult = new StockScreenerResult(quotes.ToArray());
                    if (!_industryIds.Any())
                    {
                        screenerResult.FinalResponse = true;
                    }
                }
                return screenerResult;
            }
            catch
            {
                lock (_lock)
                {
                    _industryIds.Clear(); 
                }
                var errorResult = new StockScreenerResult(quotes.ToArray());
                errorResult.FinalResponse = true;
                return errorResult;
            }
        }

        private IEnumerable<QuotesData> FilterQuotes(IEnumerable<QuotesData> quotes, IEnumerable<Criteria> criterias)
        {
            Type type = typeof(QuotesData);
            var result = quotes.ToList();
            var rangeCriterias = new List<RangeCriteria>();
            foreach (var criteria in criterias)
            {
                var rc = criteria as RangeCriteria;
                if (rc != null)
                {
                    rangeCriterias.Add(rc);
                }
            }

            rangeCriterias.RemoveAll(x => !x.IsValid);

            for (int j = 0; j < rangeCriterias.Count; j++)
            {
                result.RemoveAll(x =>
                {
                    string propertyString = type.InvokeMember(rangeCriterias[j].Name, BindingFlags.GetProperty, null, x, null) as string;
                    if (string.IsNullOrEmpty(propertyString))
                        return true;

                    double property;
                    if (rangeCriterias[j].Name.Equals("MarketCapitalization") || rangeCriterias[j].Name.Equals("EBITDA"))
                    {
                        property = DigitsParser.GetDoubleForMarCap(propertyString, 0);
                    }
                    else
                    {
                        property = DigitsParser.GetDouble(propertyString, 0);
                    }

                    if (rangeCriterias[j].Min.HasValue)
                        if (property < rangeCriterias[j].Min.Value)
                            return true;

                    if (rangeCriterias[j].Max.HasValue)
                        if (property > rangeCriterias[j].Max.Value)
                            return true;

                    return false;
                });
            }
            return result;
        }
    }
}
