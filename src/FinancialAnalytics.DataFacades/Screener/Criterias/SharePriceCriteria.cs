using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class SharePriceCriteria : RangeCriteria
    {
        public SharePriceCriteria()
        {
            Name = QuoteProperty.Open.ToString();
            DisplayName = "Share Price";
            Metrics = "$";
        }
    }
}
