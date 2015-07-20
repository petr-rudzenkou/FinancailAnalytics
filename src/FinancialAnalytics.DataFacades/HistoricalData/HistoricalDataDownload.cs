using System;
using System.Collections.Generic;
using System.Linq;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.HistoricalData
{
    public class HistoricalDataDownload : DownloadClient<HistoricalDataResult>
    {
        public Response<HistoricalDataResult> Download(string symbol, DateTime startDate, DateTime endDate)
        {
            var settings = new HistoricalDataDownloadSettings()
            {
                IDs = new[] { symbol },
                StartDate = startDate,
                EndDate = endDate
            };
            return base.Download(settings);
        }
        public void DownloadAsync(string symbol, DateTime startDate, DateTime endDate)
        {
            var settings = new HistoricalDataDownloadSettings()
            {
                IDs = new []{symbol},
                StartDate = startDate,
                EndDate = endDate
            };
            base.DownloadAsync(settings, null);
        }

        protected override HistoricalDataResult ConvertResult(ConnectionInfo connInfo, System.IO.Stream stream, SettingsBase settings)
        {
            string result = DataFacadesHelper.StreamToString(stream);
            IEnumerable<HistoricalData> historicalData = ImportExport.XmlToHistoricalData(result);
            return new HistoricalDataResult(historicalData.ToArray());
        }
    }
}
