using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Interfaces;

namespace FinancialAnalytics.DataFacades.Base
{
    public class DownloadEventArgs : EventArgs
    {
        private object mUserArgs = null;
        /// <summary>
        /// Gets the user argument that were passed when the download was started.
        /// </summary>
        /// <value></value>
        /// <returns>An object defined by the user</returns>
        /// <remarks></remarks>
        public object UserArgs
        {
            get { return mUserArgs; }
        }
        public DownloadEventArgs(object userArgs)
        {
            mUserArgs = userArgs;
        }
    }



    /// <summary>
    /// Base event class for completed asynchronous download processes that provides additionally the response of the download. This class must be inherited.
    /// </summary>
    /// <remarks></remarks>
    public class DownloadCompletedEventArgs<T> : DownloadEventArgs, IDownloadCompletedEventArgs
    {
        private SettingsBase mSettings = null;
        public SettingsBase Settings { get { return mSettings; } }

        private Response<T> mResponse = null;
        /// <summary>
        /// Gets the response of the download process.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Response<T> Response { get { return mResponse; } }
        public IResponse GetResponse() { return this.Response; }

        protected DownloadCompletedEventArgs(object userArgs, Response<T> resp, SettingsBase settings)
            : base(userArgs)
        {
            mResponse = resp;
            mSettings = settings;
        }
    }
}
