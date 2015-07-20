using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Interfaces
{
    public interface IDownload
    {
        IResponse GetResponse();
    }
}
