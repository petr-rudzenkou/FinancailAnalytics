using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class ProfitMarginCriteria : RangeCriteria
    {
        public ProfitMarginCriteria()
        {
            Name = QuoteProperty.ChangeFromYearHigh.ToString();
            DisplayName = "Profit Margin";
            Metrics = "%";
        }
    }
}
