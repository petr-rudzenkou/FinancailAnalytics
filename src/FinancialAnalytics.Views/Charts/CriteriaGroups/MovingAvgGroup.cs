using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class MovingAvgGroup : ChartCriteriaGroup
    {
        public MovingAvgGroup()
        {
            Name = "MovingAvg";
            DisplayName = "Moving Avg";
        }

        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m5,
                DisplayName = "5",
                IsSelected = false
            });
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m10,
                DisplayName = "10",
                IsSelected = false
            });
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m20,
                DisplayName = "20",
                IsSelected = false
            });
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m50,
                DisplayName = "50",
                IsSelected = false
            });
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m100,
                DisplayName = "100",
                IsSelected = false
            });
            ChartCriterias.Add(new MovingAvgCriteria()
            {
                MovingAverageInterval = MovingAverageInterval.m200,
                DisplayName = "200",
                IsSelected = false
            });
        }
    }
}
