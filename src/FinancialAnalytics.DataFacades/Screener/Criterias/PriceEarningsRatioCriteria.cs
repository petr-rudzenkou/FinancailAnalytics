using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class PriceEarningsRatioCriteria : RangeCriteria
    {
        public PriceEarningsRatioCriteria()
        {
            Name = QuoteProperty.PERatio.ToString();
            DisplayName = "Price/Earnings Ratio";
        }
    }
}
