using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotItem : IEntityWrapper<IPivotItem>
    {
        string Value { get; set; }
        int Position { get; set; }
        bool DrilledDown { get; set; }
    }
}
