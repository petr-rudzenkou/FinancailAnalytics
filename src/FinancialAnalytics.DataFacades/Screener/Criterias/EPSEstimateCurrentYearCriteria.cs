namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class EPSEstimateCurrentYearCriteria : DataFacades.Screener.Criterias.Base.RangeCriteria
    {
        public EPSEstimateCurrentYearCriteria()
        {
            Name = QuoteProperty.EPSEstimateCurrentYear.ToString();
            DisplayName = "EPS Estimate current year";
        }
    }
}
