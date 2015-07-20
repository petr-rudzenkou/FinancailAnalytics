using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.GroupContainers
{
    public class IndicatorsContainer : CriteriaGroupContainer
    {
        public IndicatorsContainer()
        {
            ChartCriteriaGroups.Add(new IndicatorsGroup());
        }
    }
}
