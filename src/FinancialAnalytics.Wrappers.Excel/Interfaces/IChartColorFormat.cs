using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IChartColorFormat
    {
        int RGB { get; }

        int SchemeColor { get; set; }

        bool Equals(IChartColorFormat obj);
    }
}
