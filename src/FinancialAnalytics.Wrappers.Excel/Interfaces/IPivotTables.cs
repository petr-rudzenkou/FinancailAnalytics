using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotTables : IEntitiesCollectionWrapper<IPivotTables, IPivotTable>
    {
        IPivotTable Add(
            IPivotCache pivotCache,
            Object tableDestination,
            Object tableName,
            Object readData,
            PivotTableVersionList defaultVersion);
    }
}
