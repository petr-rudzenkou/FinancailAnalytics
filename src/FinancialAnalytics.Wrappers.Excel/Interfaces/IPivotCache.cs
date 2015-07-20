using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotCache : IEntityWrapper<IPivotCache>
    {
        bool OLAP { get; }

        object CommandText { get; set; }

        bool IsConnected { get; }

        Object Connection { get; set; }

        bool MaintainConnection { get; set; }

        CmdType CommandType { get; set; }

        Object PivotCacheObject { get; }

        IPivotTable CreatePivotTable(
            Object tableDestination,
            Object tableName,
            Object readData,
            Object defaultVersion);
    }
}
