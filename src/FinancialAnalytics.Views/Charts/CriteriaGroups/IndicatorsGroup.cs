using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class IndicatorsGroup : ChartCriteriaGroup
    {
        public IndicatorsGroup()
        {
            Name = "Indicators";
            DisplayName = "Indicators";
        }
        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.MACD,
                DisplayName = "MACD",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.MFI,
                DisplayName = "MFI",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.ROC,
                DisplayName = "ROC",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.RSI,
                DisplayName = "RSI",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.Slow_Stoch,    
                DisplayName = "Slow Stoch",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.Fast_Stoch,
                DisplayName = "Fast Stoch",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.Vol,
                DisplayName = "Vol",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.Vol_MA,
                DisplayName = "Vol+MA",
                IsSelected = false
            });
            ChartCriterias.Add(new IndicatorCriteria()
            {
                Indicator = TechnicalIndicator.W_R,
                DisplayName = "W%R",
                IsSelected = false
            });
        }
    }
}
