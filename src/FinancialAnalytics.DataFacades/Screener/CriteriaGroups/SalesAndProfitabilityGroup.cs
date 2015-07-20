using FinancialAnalytics.DataFacades.Screener.Criterias;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    public class SalesAndProfitabilityGroup : CriteriaGroup
    {
        public SalesAndProfitabilityGroup()
        {
            DisplayName = "Sales & Profitability";
        }
        protected override void CreateCriterias()
        {
            CriteriaFilters.Add(new VolumeCriteria());
            CriteriaFilters.Add(new AverageDailyVolumeCriteria());
            CriteriaFilters.Add(new ProfitMarginCriteria());
            CriteriaFilters.Add(new EBITDACriteria());
        }
    }
}
