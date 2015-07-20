using System;
using System.Collections.Generic;
using System.Linq;

namespace FinancialAnalytics.Views.Charts.Criterias
{
    public class CompareVsCriteria : ChartCriteria
    {
        private string _ids = string.Empty;

        public List<string> CompareVsIds
        {
            get
            {
                if (!string.IsNullOrEmpty(_ids))
                {
                    return
                        _ids.Split(new[] {',', ';'}, StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .ToList();
                }
                return new List<string>();
            }
        } 
        public string Ids
        {
            get { return _ids; }
            set { _ids = value; }
        } 
    }
}
