using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.GroupContainers
{
    public class OverlaysContainer : CriteriaGroupContainer
    {
        public OverlaysContainer()
        {
            ChartCriteriaGroups.Add(new OverlaysGroup());
            ChartCriteriaGroups.Add(new CompareVsGroup());
        }
    }
}
