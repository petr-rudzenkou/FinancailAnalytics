using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class EBITDACriteria : RangeCriteria
    {
        public EBITDACriteria()
        {
            Name = QuoteProperty.EBITDA.ToString();
            DisplayName = "EBITDA";
        }
    }
}
