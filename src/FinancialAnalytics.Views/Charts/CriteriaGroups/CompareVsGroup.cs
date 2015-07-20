using System.Collections.Generic;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public class CompareVsGroup : ChartCriteriaGroup
    {
        protected override void CreateCriterias()
        {
           ChartCriterias.Add(new CompareVsCriteria()
           {
               IsSelected = true,
           });
        }
    }
}
