namespace FinancialAnalytics.DataFacades.Base
{
    public class DefaultDownloadCompletedEventArgs<T> : DownloadCompletedEventArgs<T>
    {
        internal DefaultDownloadCompletedEventArgs(object userArgs, Response<T> response, SettingsBase settings)
            : base(userArgs, response, settings)
        {
        }

        public DefaultDownloadCompletedEventArgs<N> CreateNew<N>(N newResult)
        {
            return this.CreateNew<N>(newResult, this.UserArgs);
        }
        public DefaultDownloadCompletedEventArgs<N> CreateNew<N>(N newResult, object userArgs)
        {
            return new DefaultDownloadCompletedEventArgs<N>(userArgs, new DefaultResponse<N>(this.Response.Connection, newResult), this.Settings);
        }
    }
}
