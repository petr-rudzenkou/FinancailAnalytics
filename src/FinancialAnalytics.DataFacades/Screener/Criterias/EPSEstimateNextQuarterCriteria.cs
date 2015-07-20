namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class EPSEstimateNextQuarterCriteria : DataFacades.Screener.Criterias.Base.RangeCriteria
    {
        public EPSEstimateNextQuarterCriteria()
        {
            Name = QuoteProperty.EPSEstimateNextQuarter.ToString();
            DisplayName = "EPS Estimate next quarter";
        }
    }
}
