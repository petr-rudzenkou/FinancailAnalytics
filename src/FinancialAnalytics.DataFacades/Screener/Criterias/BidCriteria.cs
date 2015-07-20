using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class BidCriteria : RangeCriteria
    {
        public BidCriteria()
        {
            Name = QuoteProperty.Bid.ToString();
            DisplayName = "Bid Price";
            Metrics = "$";
        }
    }
}
