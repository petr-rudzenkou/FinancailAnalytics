using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;

namespace FinancialAnalytics.DataFacades.Base
{
    public abstract class SettingsBase
    {
        protected virtual bool KeepAlive { get { return true; } }
        internal bool KeepAliveInternal { get { return this.KeepAlive; } }

        public abstract string GetUrl();
        public abstract object Clone();

        private List<KeyValuePair<HttpRequestHeader, string>> mAdditionalHeaders = new List<KeyValuePair<HttpRequestHeader, string>>();

        protected virtual List<KeyValuePair<HttpRequestHeader, string>> AdditionalHeaders
        {
            get
            {
                return mAdditionalHeaders;
            }
        }
        protected NameValueCollection Headers { get; set; }
        protected virtual RequestMethod Method { get { return RequestMethod.GET; } }
        protected virtual CookieContainer Cookies { get { return null; } }
        protected virtual string ContentType { get { return string.Empty; } }
        protected virtual string PostData { get { return string.Empty; } }
        protected virtual bool DownloadResponseStream { get { return true; } }

        internal string GetUrlInternal()
        {
            return this.GetUrl();
        }
        internal List<KeyValuePair<HttpRequestHeader, string>> GetAdditionalHeadersInternal { get { return mAdditionalHeaders; } }
        internal RequestMethod MethodInternal { get { return this.Method; } }
        internal CookieContainer CookiesInternal { get { return this.Cookies; } }
        internal string ContentTypeInternal { get { return this.ContentType; } }
        internal string PostDataInternal { get { return this.PostData; } }
        internal bool DownloadResponseStreamInternal { get { return this.DownloadResponseStream; } }
    }
}
