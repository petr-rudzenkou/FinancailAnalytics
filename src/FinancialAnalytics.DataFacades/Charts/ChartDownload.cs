using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Charts
{
    public class ChartDownload : DownloadClient<ChartResult>
    {
        /// <summary>
        /// Gets or sets the chart image download options.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks>By setting null/Nothing, a default instance will be setted and used for downloading.</remarks>
        public ChartDownloadSettings Settings { get { return (ChartDownloadSettings)base.Settings; } set { base.SetSettings(value); } }

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <remarks></remarks>
        public ChartDownload()
        {
            this.Settings = new ChartDownloadSettings();
        }

        /// <summary>
        /// Downloads a chart image.
        /// </summary>
        /// <param name="unmanagedID">The unmanaged ID of the stock</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public Response<ChartResult> Download(string unmanagedID)
        {
            if (unmanagedID == string.Empty)
                throw new ArgumentNullException("unmanagedID", "The passed ID is empty.");
            ChartDownloadSettings settings = (ChartDownloadSettings)this.Settings.Clone();
            settings.ID = unmanagedID;
            return this.Download(settings);
        }
        public Response<ChartResult> Download(ChartDownloadSettings settings)
        {
            return base.Download(settings);
        }

        public void DownloadAsync(string unmanagedID, object userArgs)
        {
            if (unmanagedID == string.Empty)
                throw new ArgumentNullException("unmanagedID", "The passed ID is empty.");
            ChartDownloadSettings settings = (ChartDownloadSettings)this.Settings.Clone();
            settings.ID = unmanagedID;
            base.DownloadAsync(settings, userArgs);
        }
        /// <summary>
        /// Starts an asynchronous download of an chart image.
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="userArgs">Individual user argument</param>
        public void DownloadAsync(ChartDownloadSettings settings, object userArgs)
        {
            base.DownloadAsync(settings, userArgs);
        }

        protected override ChartResult ConvertResult(ConnectionInfo connInfo, System.IO.Stream stream, SettingsBase settings)
        {
            ChartResult chartResult;
            var chartsSettings = settings as ChartDownloadSettings;
            if (chartsSettings != null)
            {
                chartResult = new ChartResult(chartsSettings.ID, DataFacadesHelper.CopyStream(stream));
            }
            else
            {
                chartResult = new ChartResult();
            }
            return chartResult;
        }
    }
}
