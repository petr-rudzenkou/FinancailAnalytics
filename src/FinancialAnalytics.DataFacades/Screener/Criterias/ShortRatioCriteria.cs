using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class ShortRatioCriteria : RangeCriteria
    {
        public ShortRatioCriteria()
        {
            Name = QuoteProperty.ShortRatio.ToString();
            DisplayName = "Price/Short Ratio";
        }
    }
}
