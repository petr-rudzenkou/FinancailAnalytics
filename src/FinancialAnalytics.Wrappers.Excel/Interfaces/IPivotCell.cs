using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotCell : IEntityWrapper<IPivotCell>
    {
        IApplication Application { get; }
        PivotCellType PivotCellType { get; }
        IPivotField PivotField { get; }
        IPivotItem PivotItem { get; }
        IPivotTable PivotTable { get; }
        IRange Range { get; }
        IPivotItemList RowItems { get; }
        IPivotItemList ColumnItems { get; }
        IPivotField DataField { get; }
    }
}
