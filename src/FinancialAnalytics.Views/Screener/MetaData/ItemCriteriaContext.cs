using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Views.Screener.MetaData
{
    public class CriteriaDataContext
    {
        public string Id { get; set; }

        public string Display { get; set; }

        public string Name { get; set; }

        public int ParentMenuId { get; set; }

        public bool MenuHasSeparator { get; set; }

        public Type DataType { get; set; }
    }
}
