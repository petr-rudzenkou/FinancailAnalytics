

using FinancialAnalytics.DataFacades.Screener.Criterias;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    public class CategoryGroup : CriteriaGroup
    {
        public CategoryGroup()
        {
            Name = "Category";
            DisplayName = "Category";
        }
        protected override void CreateCriterias()
        {
            CriteriaFilters.Add(new IndustryCriteria());
        }
    }
}
