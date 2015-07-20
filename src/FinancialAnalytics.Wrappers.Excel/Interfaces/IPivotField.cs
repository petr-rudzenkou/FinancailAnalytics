using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotField : IEntityWrapper<IPivotField>
    {
        PivotFieldOrientation Orientation { get; set; }
        IRange DataRange { get; }
        string Name { get; }
        Object VisibleItemsList { get; set; }
        ICubeField CubeField { get; }
        string CurrentPageName { get; set; }
        Object CurrentPageList { get; set; }
        string Value { get; set; }
        Object get_VisibleItems(Object index);
        IApplication Application { get; }
        object Position { get; set; }
    }
}
