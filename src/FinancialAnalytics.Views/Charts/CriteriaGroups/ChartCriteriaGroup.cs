using System.Collections.Generic;
using FinancialAnalytics.Views.Charts.Criterias;

namespace FinancialAnalytics.Views.Charts.CriteriaGroups
{
    public abstract class ChartCriteriaGroup
    {
        private readonly IList<ChartCriteria> _criterias = new List<ChartCriteria>();
        protected ChartCriteriaGroup()
        {
            CreateCriterias();
        }

        public IList<ChartCriteria> ChartCriterias
        {
            get { return _criterias; }
        }
        protected abstract void CreateCriterias();
        public string Name { get; set; }
        public string DisplayName { get; set; }
    }
}
