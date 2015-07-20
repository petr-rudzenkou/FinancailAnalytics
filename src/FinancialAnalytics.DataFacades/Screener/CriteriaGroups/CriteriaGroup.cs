using System.Collections.Generic;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.CriteriaGroups
{
    public abstract class CriteriaGroup
    {
        private readonly IList<Criteria> _criterias = new List<Criteria>();
        protected CriteriaGroup()
        {
            CreateCriterias();
        }

        public IList<Criteria> CriteriaFilters
        {
            get { return _criterias; }
        }
        protected abstract void CreateCriterias();
        public string Name { get; set; }
        public string DisplayName { get; set; }
    }
}
