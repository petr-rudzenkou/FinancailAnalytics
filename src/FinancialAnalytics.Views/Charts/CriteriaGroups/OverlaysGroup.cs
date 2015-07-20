using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    // p=b%2Cp%2Cs%2Cv
    public class OverlaysGroup : ChartCriteriaGroup
    {
        public OverlaysGroup()
        {
            Name = "Overlays";
            DisplayName = "Overlays";
        }

        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new OverlaysCriteria()
            {
                Overlay = ChartOverlay.b,
                DisplayName = "Bollinger Bands",
                IsSelected = false
            });
            ChartCriterias.Add(new OverlaysCriteria()
            {
                Overlay = ChartOverlay.p,
                DisplayName = "Parabolic SAR",
                IsSelected = false
            });
            ChartCriterias.Add(new OverlaysCriteria()
            {
                Overlay = ChartOverlay.s,
                DisplayName = "Splits",
                IsSelected = false
            });
            ChartCriterias.Add(new OverlaysCriteria()
            {
                Overlay = ChartOverlay.v,
                DisplayName = "Volume",
                IsSelected = false
            });
        }
    }
}
