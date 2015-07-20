using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace FinancialAnalytics.DataFacades.Base
{
    public class ConnectionInfo : ICloneable
    {
        private Exception mException;
        private int mTimeout = 0;
        private int mSizeInBytes = 0;
        private DateTime mStartTime;
        private DateTime mEndTime;

        private TimeSpan mTimeSpan;

        private KeyValuePair<HttpResponseHeader, string>[] mResponseHeaders = null;
        public KeyValuePair<HttpResponseHeader, string>[] ResponseHeaders { get { return mResponseHeaders; } }

        public ConnectionInfo(Exception exception, int timeout, int size, DateTime startTime, DateTime endTime, KeyValuePair<HttpResponseHeader, string>[] respHeaders)
        {
            mException = exception;
            mResponseHeaders = respHeaders;
            mTimeout = timeout;
            mSizeInBytes = size;
            mEndTime = endTime;
            mStartTime = startTime;
            mTimeSpan = mEndTime - mStartTime;
        }
        /// <summary>
        /// Indicates the connection status.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ConnectionState State
        {
            get
            {
                if (mException == null)
                {
                    return ConnectionState.Success;
                }
                else
                {
                    if (mException is System.Net.WebException)
                    {
                        System.Net.WebException exp = (System.Net.WebException)mException;
                        if (exp.Status == WebExceptionStatus.RequestCanceled) { return ConnectionState.Canceled; }
                        else if (exp.Status == TimeoutWebClient<object>.GetTimeoutStatus()) { return ConnectionState.Timeout; }
                        else { return ConnectionState.ErrorOccured; }
                    }
                    else
                    {
                        return ConnectionState.ErrorOccured;
                    }

                }
            }
        }
        /// <summary>
        /// If an exception occurs during download process, the exception object will be stored here. If no exception occurs, this property is null/Nothing.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Exception Exception
        {
            get { return mException; }
        }
        /// <summary>
        /// The size of downloaded data in bytes.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int SizeInBytes
        {
            get { return mSizeInBytes; }
        }
        /// <summary>
        /// The setted timeout for download process in milliseconds.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int Timeout
        {
            get { return mTimeout; }
        }
        /// <summary>
        /// The start time of download process, independent to individual preparation of passed parameters for start downloading.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public DateTime StartTime
        {
            get { return mStartTime; }
        }
        /// <summary>
        /// The end time of the download process, independent to post-processing actions like parsing.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public DateTime EndTime
        {
            get { return mEndTime; }
        }
        /// <summary>
        /// The time span of start and end time.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public TimeSpan TimeSpan
        {
            get { return mTimeSpan; }
        }
        public double KBPerSecond
        {
            get
            {
                if (this.TimeSpan.TotalMilliseconds != 0)
                {
                    return mSizeInBytes / this.TimeSpan.TotalMilliseconds;
                }
                else
                {
                    return 0;
                }
            }
        }


        public ConnectionInfo(Exception exception, int timeout, int size, DateTime startTime, DateTime endTime)
        {
            mException = exception;
            mTimeout = timeout;
            mSizeInBytes = size;
            mEndTime = endTime;
            mStartTime = startTime;
            mTimeSpan = mEndTime - mStartTime;
        }

        public virtual object Clone()
        {
            return new ConnectionInfo(this.Exception, this.Timeout, this.SizeInBytes, this.StartTime, this.EndTime);
        }
    }
}
