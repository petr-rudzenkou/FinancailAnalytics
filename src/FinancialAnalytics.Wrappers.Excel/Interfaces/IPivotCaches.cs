using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPivotCaches : IEntitiesCollectionWrapper<IPivotCaches, IPivotCache>
    {
        IPivotCache Create(PivotTableSourceType sourceType, Object sourceData, Object version);

        IPivotCache Add(PivotTableSourceType sourceType, Object sourceData);
    }
}
