using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class SalesRevenueCriteria : RangeCriteria
    {
        public SalesRevenueCriteria()
        {
            Name = QuoteProperty.BookValue.ToString();
            DisplayName = "Book Value";
            Metrics = "Mil";
        }
    }
}
