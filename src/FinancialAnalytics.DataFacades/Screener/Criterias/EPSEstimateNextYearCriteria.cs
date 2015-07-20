namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class EPSEstimateNextYearCriteria : DataFacades.Screener.Criterias.Base.RangeCriteria
    {
        public EPSEstimateNextYearCriteria()
        {
            Name = QuoteProperty.EPSEstimateNextYear.ToString();
            DisplayName = "EPS Estimate next year";
        }
    }
}
