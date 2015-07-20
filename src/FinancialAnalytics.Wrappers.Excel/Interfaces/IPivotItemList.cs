using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotItemList : IEntityWrapper<IPivotItemList>, IEnumerable
    {
        int Count { get; }
        IPivotItem this[int index] { get; }
    }
}
