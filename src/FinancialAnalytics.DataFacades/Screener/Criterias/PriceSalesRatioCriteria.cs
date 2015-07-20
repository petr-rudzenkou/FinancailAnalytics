using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class PriceSalesRatioCriteria : RangeCriteria
    {
        public PriceSalesRatioCriteria()
        {
            Name = QuoteProperty.PriceSales.ToString();
            DisplayName = "Price/Sales Ratio";
        }
    }
}
