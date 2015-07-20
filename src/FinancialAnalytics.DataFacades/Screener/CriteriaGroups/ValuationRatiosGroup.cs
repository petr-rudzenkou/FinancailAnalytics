using FinancialAnalytics.DataFacades.Screener.Criterias;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    public class ValuationRatiosGroup : CriteriaGroup
    {
        public ValuationRatiosGroup()
        {
            DisplayName = "Valuation Ratios";
        }
        protected override void CreateCriterias()
        {
            CriteriaFilters.Add(new PriceEarningsRatioCriteria());
            CriteriaFilters.Add(new PriceBookRatioCriteria());
            CriteriaFilters.Add(new PriceSalesRatioCriteria());
            CriteriaFilters.Add(new PEGRatioCriteria());
            CriteriaFilters.Add(new ShortRatioCriteria());
        }
    }
}
