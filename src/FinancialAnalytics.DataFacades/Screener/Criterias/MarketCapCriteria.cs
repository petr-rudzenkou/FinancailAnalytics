using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class MarketCapCriteria : RangeCriteria
    {
        public MarketCapCriteria()
        {
            Name = QuoteProperty.MarketCapitalization.ToString();
            DisplayName = "Market Cap";
            Metrics = "Mil";
        }
    }
}
