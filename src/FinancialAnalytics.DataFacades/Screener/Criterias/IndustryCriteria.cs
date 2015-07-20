using System.Collections.Generic;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class IndustryCriteria : Criteria
    {
        private DataFacades.Screener.Metedata.Industry _selectedIndustry;
        public IEnumerable<DataFacades.Screener.Metedata.Industry> Industries { get; set; }

        public IndustryCriteria()
        {
            Name = "Industry";
            DisplayName = "Industry";
            Industries = new List<Metedata.Industry>(DataFacades.Screener.Metedata.Industries.All);
        }
        public DataFacades.Screener.Metedata.Industry SelectedIndustry
        {
            get { return _selectedIndustry; }
            set
            {
                _selectedIndustry = value;
            }
        }
    }
}
