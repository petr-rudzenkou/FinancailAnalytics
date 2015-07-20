using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Quotes
{
    public class QuotesDownload : DownloadClient<QuotesResult>
    {

        public QuotesDownloadSettings Settings { get { return (QuotesDownloadSettings)base.Settings; } set { base.SetSettings(value); } }

        public QuotesDownload()
        {
            this.Settings = new QuotesDownloadSettings();
        }

        public Response<QuotesResult> Download(IEnumerable<string> unmanagedIDs)
        {
            return this.Download(unmanagedIDs, Settings.Properties);
        }
       
        /// <summary>
        /// Downloads quotes data.
        /// </summary>
        /// <param name="unmanagedID">The unmanaged ID</param>
        /// <param name="properties">The properties of each quote data. If parameter is null/Nothing, Symbol and LastTradePrizeOnly will set as property. In this case, with YQL server you will get every available property.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public Response<QuotesResult> Download(string unmanagedID, IEnumerable<QuoteProperty> properties)
        {
            if (unmanagedID == string.Empty)
                throw new ArgumentNullException("unmanagedID", "The passed id is empty.");
            return this.Download(new string[] { unmanagedID }, properties);
        }

        /// <summary>
        /// Downloads quotes data.
        /// </summary>
        /// <param name="unmanagedIDs">The list of unmanaged IDs</param>
        /// <param name="properties">The properties of each quote data. If parameter is null/Nothing, Symbol and LastTradePrizeOnly will set as property. In this case, with YQL server you will get every available property.</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public Response<QuotesResult> Download(IEnumerable<string> unmanagedIDs, IEnumerable<QuoteProperty> properties)
        {
            if (unmanagedIDs == null)
                throw new ArgumentNullException("unmanagedIDs", "The passed list is null.");
            return this.Download(new QuotesDownloadSettings() { IDs = unmanagedIDs.ToArray(), Properties = DataFacadesHelper.EnumToArray(properties) });
        }
        public Response<QuotesResult> Download(QuotesDownloadSettings settings)
        {
            return base.Download(settings);
        }


        public void DownloadAsync(IEnumerable<string> unmanagedIDs, object userArgs)
        {
            this.DownloadAsync(unmanagedIDs, Settings.Properties, userArgs);
        }

        /// <summary>
        /// Starts an asynchronous download of quotes data.
        /// </summary>
        /// <param name="unmanagedID">The unmanaged ID</param>
        /// <param name="properties">The properties of each quote data. If parameter is null/Nothing, Symbol and LastTradePrizeOnly will set as property. In this case, with YQL server you will get every available property.</param>
        /// <param name="userArgs">Individual user argument</param>
        /// <remarks></remarks>
        public void DownloadAsync(string unmanagedID, IEnumerable<QuoteProperty> properties, object userArgs)
        {
            if (unmanagedID == string.Empty)
                throw new ArgumentNullException("unmanagedID", "The passed ID is empty.");
            this.DownloadAsync(new string[] { unmanagedID }, properties, userArgs);
        }
        
        /// <summary>
        /// Starts an asynchronous download of quotes data.
        /// </summary>
        /// <param name="unmanagedIDs">The list of unmanaged IDs</param>
        /// <param name="properties">The properties of each quote data. If parameter is null/Nothing, Symbol and LastTradePrizeOnly will set as property. In this case, with YQL server you will get every available property.</param>
        /// <param name="userArgs">Individual user argument</param>
        /// <remarks></remarks>
        public void DownloadAsync(IEnumerable<string> unmanagedIDs, IEnumerable<QuoteProperty> properties, object userArgs)
        {
            this.DownloadAsync(new QuotesDownloadSettings() { IDs = unmanagedIDs.ToArray(), Properties = DataFacadesHelper.EnumToArray(properties) }, userArgs);
        }

        /// <summary>
        /// Starts an asynchronous download of quotes data.
        /// </summary>
        /// <param name="settings">Individual Download Settings.</param>
        /// <param name="userArgs">Individual user argument.</param>
        public void DownloadAsync(QuotesDownloadSettings settings, object userArgs)
        {
            base.DownloadAsync(settings, userArgs);
        }

        protected override QuotesResult ConvertResult(ConnectionInfo connInfo, Stream stream, SettingsBase settings)
        {
            string result = DataFacadesHelper.StreamToString(stream);
            IEnumerable<QuotesData> quotesData = ImportExport.XmlToQuoteData(result);
            return new QuotesResult(quotesData.ToArray());
        }
    }
}
