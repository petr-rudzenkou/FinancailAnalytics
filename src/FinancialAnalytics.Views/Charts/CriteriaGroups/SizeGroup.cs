using FinancialAnalytics.DataFacades;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class SizeGroup : ChartCriteriaGroup
    {
        public SizeGroup()
        {
            Name = "Size";
            DisplayName = "Size";
        }

        protected override void CreateCriterias()
        {
            //ChartCriterias.Add(new SizeCriteria()
            //{
            //    Size = ChartImageSize.Small,
            //    DisplayName = "S",
            //    IsSelected = false
            //});
            ChartCriterias.Add(new SizeCriteria()
            {
                Size = ChartImageSize.Middle,
                DisplayName = "M",
                IsSelected = false
            });
            ChartCriterias.Add(new SizeCriteria()
            {
                Size = ChartImageSize.Large,
                DisplayName = "L",
                IsSelected = true
            });
        }
    }
}
