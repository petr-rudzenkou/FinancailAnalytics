using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class ScaleGroup : ChartCriteriaGroup
    {
        public ScaleGroup()
        {
            Name = "Scale";
            DisplayName = "Scale";
        }

        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new ScaleCriteria()
            {
                Scale = ChartScale.Arithmetic,
                DisplayName = "Linear",
                IsSelected = false
            });
            ChartCriterias.Add(new ScaleCriteria()
            {
                Scale = ChartScale.Logarithmic,
                DisplayName = "Log",
                IsSelected = true
            });
        }
    }
}
