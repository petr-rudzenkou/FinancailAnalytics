namespace FinancialAnalytics.DataFacades.Screener.Criterias.Base
{
    public class RangeCriteria : Criteria
    {
        public double? Max { get; set; }
        public double? Min { get; set; }
        public string Metrics { get; set; }

        public bool IsValid
        {
            get { return Min.HasValue || Max.HasValue; }
        }
    }
}
