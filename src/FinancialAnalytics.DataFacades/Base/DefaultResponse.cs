namespace FinancialAnalytics.DataFacades.Base
{
    public class DefaultResponse<T> : Response<T>
    {
        internal DefaultResponse(ConnectionInfo info, T result)
            : base(info, result)
        {
        }
        public DefaultResponse<N> CreateNew<N>(N newResult)
        {
            return new DefaultResponse<N>(this.Connection, newResult);
        }
    }
}