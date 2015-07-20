using System.Collections.Generic;
using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.GroupContainers
{
    public class CriteriaGroupContainer
    {
        private readonly List<ChartCriteriaGroup>_chartCriteriaGroups = new List<ChartCriteriaGroup>();
        public List<ChartCriteriaGroup> ChartCriteriaGroups
        {
            get { return _chartCriteriaGroups; }
        }
    }
}
