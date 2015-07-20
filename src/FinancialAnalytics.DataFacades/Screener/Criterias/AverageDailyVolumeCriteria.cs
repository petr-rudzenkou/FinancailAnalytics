using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener.Criterias
{
    public class AverageDailyVolumeCriteria : RangeCriteria
    {
        public AverageDailyVolumeCriteria()
        {
            Name = QuoteProperty.AverageDailyVolume.ToString();
            DisplayName = "AVG Volume";
            Metrics = "K";
        }
    }
}
