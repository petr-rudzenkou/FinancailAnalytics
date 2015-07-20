using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class PEGRatioCriteria : RangeCriteria
    {
        public PEGRatioCriteria()
        {
            Name = QuoteProperty.PEGRatio.ToString();
            DisplayName = "PEG Ratio";
        }
    }
}
