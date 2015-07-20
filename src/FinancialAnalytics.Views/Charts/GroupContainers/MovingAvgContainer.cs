using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.GroupContainers
{
    public class MovingAvgContainer : CriteriaGroupContainer
    {
        public MovingAvgContainer()
        {
            ChartCriteriaGroups.Add(new MovingAvgGroup());
            ChartCriteriaGroups.Add(new EMAGroup());
        }
    }
}
