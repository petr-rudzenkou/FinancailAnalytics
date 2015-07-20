using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWorkbookConnection : IEntityWrapper<IWorkbookConnection>
    {
        string Name { get; set; }
        Object WorkbookConnectionObject { get; }
    }
}
