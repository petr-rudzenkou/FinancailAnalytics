using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class DividendYieldCriteria : RangeCriteria
    {
        public DividendYieldCriteria()
        {
            Name = QuoteProperty.DividendYield.ToString();
            DisplayName = "Dividend Yield";
            Metrics = "%";
        }
    }
}
