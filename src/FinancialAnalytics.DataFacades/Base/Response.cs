using FinancialAnalytics.DataFacades.Interfaces;

namespace FinancialAnalytics.DataFacades.Base
{
    public class Response<T> : IResponse
    {
        private ConnectionInfo mConnection = null;
        private T mResult;

        /// <summary>
        /// Gets connection information of the download process.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ConnectionInfo Connection { get { return mConnection; } }
        /// <summary>
        /// Gets the received managed data.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public T Result { get { return mResult; } }
        public object GetObjectResult() { return this.Result; }

        protected Response(ConnectionInfo connInfo, T result)
        {
            mConnection = connInfo;
            mResult = result;
        }
    }
}
