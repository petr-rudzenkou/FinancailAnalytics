using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.XChangeRates.Metadata;

namespace FinancialAnalytics.DataFacades.XChangeRates
{
    public class XChangeRatesDownload : DownloadClient<XChangeRatesResult>
    {
        public Response<XChangeRatesResult> Download(IEnumerable<string> ids)
        {
            var settings = new XChangeRatesDownloadSettings()
            {
                IDs = ids.ToArray()
            };
            return base.Download(settings);
        }
        public void DownloadAsync(IEnumerable<string> ids)
        {
            var settings = new XChangeRatesDownloadSettings()
            {
                IDs = ids.ToArray()
            };
            base.DownloadAsync(settings, null);
        }
        protected override XChangeRatesResult ConvertResult(ConnectionInfo connInfo, Stream stream, SettingsBase settings)
        {
            var rates = new List<XChangeRate>();
            try
            {
                string result = DataFacadesHelper.StreamToString(stream);
                rates.AddRange(ImportExport.XmlToXChangeRate(result));
            }
            catch (Exception)
            {
                //log
            }

            return new XChangeRatesResult(rates.ToArray());
        }
    }
}
