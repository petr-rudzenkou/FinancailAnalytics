using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using FinancialAnalytics.DataFacades.Base;
using SettingsBase = FinancialAnalytics.DataFacades.Base.SettingsBase;

namespace FinancialAnalytics.DataFacades
{
    public class TimeoutWebClient<T> : IDisposable
    {
        private IWebProxy mProxy = WebRequest.DefaultWebProxy;
        public IWebProxy Proxy
        {
            get { return mProxy; }
            set { mProxy = value; }
        }


        public TimeoutWebClient(int timeout)
        {
            this.Timeout = timeout;
        }
        private void AddHeaders(HttpWebRequest wr, List<KeyValuePair<HttpRequestHeader, string>> headers)
        {
            foreach (KeyValuePair<HttpRequestHeader, String> header in headers)
            {
                wr.Headers.Add(header.Key, header.Value);
            }
        }

        public Response<System.IO.Stream> Download(SettingsBase userSettings)
        {
            if (!mDisposedValue)
            {

                StreamDownloadSettings<T> ss = new StreamDownloadSettings<T>(userSettings);
                DateTime startTime = System.DateTime.Now;
                HttpWebRequest wr = this.GetWebRequest(ss);
                byte[] postDataBytes = null;
                if (userSettings.PostDataInternal != string.Empty)
                {
                    postDataBytes = System.Text.Encoding.ASCII.GetBytes(userSettings.PostDataInternal);
                    wr.ContentLength = postDataBytes.Length;
                }
                mActualDownload = wr;

                System.IO.MemoryStream memStream = null;
                System.Net.WebException dlException = null;
                int size = 0;
                List<KeyValuePair<HttpResponseHeader, string>> headers = new List<KeyValuePair<HttpResponseHeader, string>>();
                DateTime endTime = System.DateTime.Now;

                try
                {
                    if (postDataBytes != null)
                    {
                        using (System.IO.Stream s = wr.GetRequestStream())
                        {
                            s.Write(postDataBytes, 0, postDataBytes.Length);
                        }
                    }

                    using (HttpWebResponse resp = (HttpWebResponse)wr.GetResponse())
                    {
                        foreach (var header in resp.Headers.Keys)
                        {
                            headers.Add(new KeyValuePair<HttpResponseHeader, string>());
                        }
                        if (userSettings.DownloadResponseStreamInternal)
                        {
                            System.IO.Stream s = resp.GetResponseStream();
                            endTime = System.DateTime.Now;
                            memStream = DataFacadesHelper.CopyStream(s);
                            s.Dispose();
                        }
                    }

                    if (memStream != null && memStream.CanSeek)
                        int.TryParse(memStream.Length.ToString(), out size);
                }
                catch (Exception ex)
                {
                    dlException = this.GetOrCreateWebException(ex, null);
                }
                finally
                {
                    mActualDownload = null;
                }


                return new DefaultResponse<System.IO.Stream>(new ConnectionInfo(dlException, this.Timeout, size, startTime, endTime, headers.ToArray()), memStream);

            }
            else
            {
                return null;
            }
        }


        private void SetProxyAndTimeout(HttpWebRequest wr, SettingsBase s)
        {
            if (mProxy != null)
                wr.Proxy = mProxy;
            wr.Timeout = this.Timeout;
            wr.KeepAlive = s.KeepAliveInternal;
        }
        private byte[] StringToAscii(string s)
        {
            return System.Text.Encoding.ASCII.GetBytes(s);
        }
        public static WebExceptionStatus GetTimeoutStatus()
        {
            return WebExceptionStatus.Timeout;
        }

        public event AsyncDownloadCompletedEventHandler AsyncDownloadCompleted;
        public delegate void AsyncDownloadCompletedEventHandler(TimeoutWebClient<T> sender, StreamDownloadCompletedEventArgs<T> e);

        private delegate void AsyncCompletedDelegate(StreamDownloadCompletedEventArgs<T> e);

        private object mUserArgs = null;
        public object UserArgs
        {
            get { return mUserArgs; }
        }

        private HttpWebRequest mActualDownload = null;
        public HttpWebRequest ActualDownload
        {
            get { return mActualDownload; }
        }

        public bool IsBusy
        {
            get { return this.ActualDownload != null; }
        }

        public int Timeout { get; set; }

        public void CancelAsync()
        {
            try
            {
                if (this.ActualDownload != null)
                    this.ActualDownload.Abort();
            }
            catch (ArgumentOutOfRangeException ex)
            {
                this.Dispose();
            }
        }


        private HttpWebRequest GetWebRequest(StreamDownloadSettings<T> ss)
        {
            string url = ss.UserSettings.GetUrlInternal();
            HttpWebRequest wr = (HttpWebRequest)HttpWebRequest.Create(url);
            this.SetProxyAndTimeout(wr, ss.UserSettings);
            wr.Method = ss.UserSettings.MethodInternal.ToString();
            if (ss.UserSettings.CookiesInternal != null) wr.CookieContainer = ss.UserSettings.CookiesInternal;
            if (ss.UserSettings.ContentTypeInternal != string.Empty) wr.ContentType = ss.UserSettings.ContentTypeInternal;
            this.AddHeaders(wr, ss.UserSettings.GetAdditionalHeadersInternal);
            return wr;
        }

        public void DownloadAsync(SettingsBase userSettings, object userArgs)
        {
            if (!mDisposedValue)
            {
                mUserArgs = userArgs;
                StreamDownloadSettings<T> ss = new StreamDownloadSettings<T>(userSettings);
                DateTime startTime = System.DateTime.Now;
                HttpWebRequest wr = this.GetWebRequest(ss);
                byte[] postDataBytes = null;
                if (userSettings.PostDataInternal != string.Empty)
                {
                    postDataBytes = this.StringToAscii(userSettings.PostDataInternal);
                }
                mActualDownload = wr;

                try
                {
                    if (postDataBytes != null)
                    {
                        AsyncDownloadArgs<T> asyncArgs = new AsyncDownloadArgs<T>(userArgs, startTime, wr, false, ss, postDataBytes) { Timeout = this.Timeout };
                        IAsyncResult res = wr.BeginGetRequestStream(new AsyncCallback(this.RequestStreamDownloadCompleted), asyncArgs);
                        try
                        {
                            System.Threading.ThreadPool.RegisterWaitForSingleObject(res.AsyncWaitHandle, new System.Threading.WaitOrTimerCallback(this.ResponseDownloadTimeout), asyncArgs, (this.Timeout), true);
                        }
                        catch (NotSupportedException ex)
                        {
                        }
                    }
                    else
                    {
                        this.DownloadAsyncResponse(wr, ss, startTime, userArgs);
                    }

                }
                catch (Exception ex)
                {
                    System.Net.WebException dlException = this.GetOrCreateWebException(ex, null);
                    if (AsyncDownloadCompleted != null)
                    {
                        this.AsyncDownloadCompleted(this, new StreamDownloadCompletedEventArgs<T>(userArgs, new DefaultResponse<Stream>(new ConnectionInfo(dlException, this.Timeout, 0, startTime, System.DateTime.Now), null), ss));
                    }
                }
            }
        }

        private void RequestStreamDownloadCompleted(IAsyncResult result)
        {
            AsyncDownloadArgs<T> asyncArgs = (AsyncDownloadArgs<T>)result.AsyncState;
            using (Stream s = asyncArgs.WR.EndGetRequestStream(result))
            {
                s.Write(asyncArgs.PostData, 0, asyncArgs.PostData.Length);
            }
            this.DownloadAsyncResponse(asyncArgs.WR, asyncArgs.Settings, asyncArgs.StartTime, asyncArgs.UserArgs);
        }


        private void DownloadAsyncResponse(HttpWebRequest wr, StreamDownloadSettings<T> ss, DateTime startTime, object userArgs)
        {
            try
            {
                AsyncDownloadArgs<T> asyncArgs = new AsyncDownloadArgs<T>(userArgs, System.DateTime.Now, wr, false, ss, null);

                IAsyncResult res = wr.BeginGetResponse(new AsyncCallback(this.ResponseDownloadCompleted), asyncArgs);
                try
                {
                    //System.Threading.ThreadPool.RegisterWaitForSingleObject(res.AsyncWaitHandle, new System.Threading.WaitOrTimerCallback(this.ResponseDownloadTimeout), asyncArgs, (this.Timeout), true);
                }
                catch (NotSupportedException ex)
                {
                }
            }
            catch (Exception ex)
            {
                System.Net.WebException dlException = this.GetOrCreateWebException(ex, null);
                if (AsyncDownloadCompleted != null)
                {
                    this.AsyncDownloadCompleted(this, new StreamDownloadCompletedEventArgs<T>(userArgs, new DefaultResponse<Stream>(new ConnectionInfo(dlException, this.Timeout, 0, startTime, System.DateTime.Now), null), ss));
                }
            }
        }




        private void ResponseDownloadTimeout(object state, bool timedOut)
        {
            if (timedOut)
            {
                AsyncDownloadArgs<T> asyncArgs = (AsyncDownloadArgs<T>)state;
                asyncArgs.TimedOut = true;
                asyncArgs.WR.Abort();
            }
        }
        private void ResponseDownloadCompleted(IAsyncResult result)
        {
            if (!mDisposedValue)
            {
                mActualDownload = null;
                AsyncDownloadArgs<T> asyncArgs = (AsyncDownloadArgs<T>)result.AsyncState;
                Stream memStream = null;
                WebException dlException = null;
                System.DateTime endTime = System.DateTime.Now;
                int size = 0;

                try
                {
                    HttpWebResponse resp = (HttpWebResponse)asyncArgs.WR.EndGetResponse(result);

                    if (!asyncArgs.TimedOut)
                    {
                        try
                        {
                            if (asyncArgs.Settings.DownloadResponseStreamInternal)
                            {
                                memStream = DataFacadesHelper.CopyStream(resp.GetResponseStream());
                                endTime = DateTime.Now;
                            }
                        }
                        catch (Exception ex)
                        {
                            dlException = this.GetOrCreateWebException(ex, resp);
                        }
                        finally
                        {
                            resp.Close();
                        }
                    }
                    else
                    {
                        dlException = new WebException("Timeout Exception.", null, GetTimeoutStatus(), resp);
                    }

                }
                catch (Exception ex)
                {
                    dlException = this.GetOrCreateWebException(ex, null);
                }
                finally
                {
                    if (memStream != null && memStream.CanSeek)
                        int.TryParse(memStream.Length.ToString(), out size);
                    mUserArgs = null;

                    StreamDownloadCompletedEventArgs<T> e = new StreamDownloadCompletedEventArgs<T>(asyncArgs.UserArgs, new DefaultResponse<Stream>(new ConnectionInfo(dlException, asyncArgs.Timeout, size, asyncArgs.StartTime, endTime), memStream), asyncArgs.Settings);
                    if (this.AsyncDownloadCompleted != null) AsyncDownloadCompleted(this, e);
                }
            }
        }

        private WebException GetOrCreateWebException(Exception ex, WebResponse resp)
        {
            if (ex is WebException)
            {
                return (WebException)ex;
            }
            else
            {
                return new System.Net.WebException("A Non-Net.WebException was throwed during download process. See [InnerException] for more details.", ex, System.Net.WebExceptionStatus.UnknownError, resp);
            }
        }

        #region "IDisposable"
        private bool mDisposedValue = false;
        protected virtual void Dispose(bool disposing)
        {
            if (!mDisposedValue)
            {
                mDisposedValue = true;
                if (disposing)
                {
                }
                mUserArgs = null;
                mActualDownload = null;
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }

    internal class StreamDownloadSettings<T> : SettingsBase
    {

        private SettingsBase mUserSettings = null;
        public SettingsBase UserSettings { get { return mUserSettings; } }
        internal StreamDownloadSettings(SettingsBase settings)
        {
            mUserSettings = settings;
        }
        public override string GetUrl()
        {
            return mUserSettings.GetUrlInternal();
        }
        public override object Clone()
        {
            return new StreamDownloadSettings<T>(this.UserSettings);
        }
    }

    public class StreamDownloadCompletedEventArgs<T> : DownloadCompletedEventArgs<Stream>
    {
        public SettingsBase UserSettings { get { return ((StreamDownloadSettings<T>)base.Settings).UserSettings; } }
        internal StreamDownloadCompletedEventArgs(object userArgs, Response<Stream> response, StreamDownloadSettings<T> settings)
            : base(userArgs, response, settings)
        {
        }

    }


    internal class AsyncDownloadArgs<R> : DownloadEventArgs
    {
        public DateTime StartTime;
        public HttpWebRequest WR = null;
        public int Timeout = 0;
        public bool TimedOut = false;
        public StreamDownloadSettings<R> Settings = null;
        public byte[] PostData = null;
        public AsyncDownloadArgs(object userArgs, DateTime st, HttpWebRequest wr, bool isTime, StreamDownloadSettings<R> settings, byte[] postData)
            : base(userArgs)
        {
            this.StartTime = st;
            this.WR = wr;
            this.TimedOut = isTime;
            this.Settings = settings;
            this.PostData = postData;
        }
    }
}
