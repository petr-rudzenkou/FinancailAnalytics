using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IListRows
    {
        int Count { get; }

        bool Equals(IListRows obj);
    }
}
