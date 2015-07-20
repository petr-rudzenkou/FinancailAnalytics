using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class VolumeCriteria : RangeCriteria
    {
        public VolumeCriteria()
        {
            Name = QuoteProperty.Volume.ToString();
            DisplayName = "Volume";
            Metrics = "K";
        }
    }
}
