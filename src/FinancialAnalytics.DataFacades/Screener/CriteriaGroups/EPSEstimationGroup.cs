using System;
using FinancialAnalytics.DataFacades.Screener.Criterias;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    // Currently not used
    [Obsolete]
    public class EPSEstimationGroup : CriteriaGroup
    {
        public EPSEstimationGroup()
        {
            DisplayName = "EPS Estiamation";
        }

        protected override void CreateCriterias()
        {
            CriteriaFilters.Add(new EPSEstimateCurrentYearCriteria());
            CriteriaFilters.Add(new EPSEstimateNextYearCriteria());
            CriteriaFilters.Add(new EPSEstimateNextQuarterCriteria());
        }
    }
}
