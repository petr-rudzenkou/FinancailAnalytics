using FinancialAnalytics.DataFacades.Screener.Criterias;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    public class ShareDataGroup : CriteriaGroup
    {
        public ShareDataGroup()
        {
            DisplayName = "Share Data";
        }
        protected override void CreateCriterias()
        {
            CriteriaFilters.Add(new SharePriceCriteria());
            CriteriaFilters.Add(new MarketCapCriteria());
            CriteriaFilters.Add(new DividendYieldCriteria());
            CriteriaFilters.Add(new AskCriteria());
            CriteriaFilters.Add(new BidCriteria());
        }
    }
}
