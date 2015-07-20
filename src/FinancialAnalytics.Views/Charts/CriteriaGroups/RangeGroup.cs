using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class RangeGroup : ChartCriteriaGroup
    {
        private RangeCriteria _selectedRangeCriteria;
        public RangeGroup()
        {
            Name = "Range";
            DisplayName = "Range";
        }

        public RangeCriteria SelectedChartCriteria
        {
            get { return _selectedRangeCriteria; }
            set
            {
                _selectedRangeCriteria = value;
            }
        }

        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c1D,
                DisplayName = "1d",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c5D,
                DisplayName = "5d",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c1M,
                DisplayName = "1m",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c3M,
                DisplayName = "3m",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c6M,
                DisplayName = "6m",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c1Y,
                DisplayName = "1y",
                IsSelected = true
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c2Y,
                DisplayName = "2y",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.c5Y,
                DisplayName = "5y",
                IsSelected = false
            });
            ChartCriterias.Add(new RangeCriteria()
            {
                Range = ChartTimeSpan.cMax,
                DisplayName = "max",
                IsSelected = false
            });
        }
    }
}
