using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class TypeGroup : ChartCriteriaGroup
    {
        public TypeGroup()
        {
            Name = "Type";
            DisplayName = "Type";
        }

        protected override void CreateCriterias()
        {
            ChartCriterias.Add(new TypeCriteria()
            {
                Type = ChartType.Bar,
                DisplayName = "Bar",
                IsSelected = false
            });
            ChartCriterias.Add(new TypeCriteria()
            {
                Type = ChartType.Line,
                DisplayName = "Line",
                IsSelected = true
            });
            ChartCriterias.Add(new TypeCriteria()
            {
                Type = ChartType.Candle,
                DisplayName = "Candle",
                IsSelected = false
            });
        }
    }
}
