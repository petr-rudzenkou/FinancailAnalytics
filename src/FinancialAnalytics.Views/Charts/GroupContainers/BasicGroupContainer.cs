using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.GroupContainers
{
    public class BasicGroupContainer : CriteriaGroupContainer
    {
        public BasicGroupContainer()
        {
            ChartCriteriaGroups.Add(new RangeGroup());
            ChartCriteriaGroups.Add(new TypeGroup());
            ChartCriteriaGroups.Add(new ScaleGroup());
            ChartCriteriaGroups.Add(new SizeGroup());
        }
    }
}
