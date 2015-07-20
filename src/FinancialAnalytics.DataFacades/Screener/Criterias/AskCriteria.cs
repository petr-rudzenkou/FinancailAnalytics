using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class AskCriteria : RangeCriteria
    {
        public AskCriteria()
        {
            Name = QuoteProperty.Ask.ToString();
            DisplayName = "Ask Price";
            Metrics = "$";
        }
    }
}
