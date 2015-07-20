using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotTable : IEntityWrapper<IPivotTable>
    {
        IPivotCache PivotCache();
        Object get_DataFields(Object index);
        string Name { get; }
        PivotTableVersionList Version { get; }
        IApplication Application { get; }
        bool EnableDrilldown { get; set; }
        void Update();
        IRange TableRange2 { get; }
        bool ShowTableStyleRowStripes { get; set; }
        Object TableStyle2 { get; set; }
        bool ManualUpdate { get; set; }
        Object Parent { get; }
        IWorksheet ParentWorksheet { get; }
        IRange RowRange { get; }
        Object PivotFields(Object index);
        Object get_PageFields(Object index);
        ICubeFields CubeFields { get; }
        Object get_ColumnFields(Object index);
        Object get_RowFields(Object index);
        IRange TableRange1 { get; }
    }
}
