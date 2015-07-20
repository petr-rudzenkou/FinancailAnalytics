using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class EMAGroup : ChartCriteriaGroup
    {
        public EMAGroup()
        {
            Name = "EMA";
            DisplayName = "EMA";
        }
        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m5,
                DisplayName = "5",
                IsSelected = false
            });
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m10,
                DisplayName = "10",
                IsSelected = false
            });
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m20,
                DisplayName = "20",
                IsSelected = false
            });
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m50,
                DisplayName = "50",
                IsSelected = false
            });
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m100,
                DisplayName = "100",
                IsSelected = false
            });
            ChartCriterias.Add(new EMACriteria()
            {
                EMA = MovingAverageInterval.m200,
                DisplayName = "200",
                IsSelected = false
            });
        }
    }
}
