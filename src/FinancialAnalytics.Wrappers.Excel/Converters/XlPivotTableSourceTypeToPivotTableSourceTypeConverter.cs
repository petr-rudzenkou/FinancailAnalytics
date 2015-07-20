using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPivotTableSourceTypeToPivotTableSourceTypeConverter
    {
        public static PivotTableSourceType Convert(XlPivotTableSourceType xlSourceType)
        {
            PivotTableSourceType result = PivotTableSourceType.PivotTable;
            switch (xlSourceType)
            {
                case XlPivotTableSourceType.xlConsolidation:
                    result = PivotTableSourceType.Consolidation;
                    break;
                case XlPivotTableSourceType.xlDatabase:
                    result = PivotTableSourceType.Database;
                    break;
                case XlPivotTableSourceType.xlExternal:
                    result = PivotTableSourceType.External;
                    break;
                case XlPivotTableSourceType.xlScenario:
                    result = PivotTableSourceType.Scenario;
                    break;
            }
            return result;
        }

        public static XlPivotTableSourceType ConvertBack(PivotTableSourceType sourceType)
        {
            XlPivotTableSourceType result = XlPivotTableSourceType.xlPivotTable;
            switch (sourceType)
            {
                case PivotTableSourceType.Consolidation:
                    result = XlPivotTableSourceType.xlConsolidation;
                    break;
                case PivotTableSourceType.Database:
                    result = XlPivotTableSourceType.xlDatabase;
                    break;
                case PivotTableSourceType.External:
                    result = XlPivotTableSourceType.xlExternal;
                    break;
                case PivotTableSourceType.PivotTable:
                    result = XlPivotTableSourceType.xlPivotTable;
                    break;
                case PivotTableSourceType.Scenario:
                    result = XlPivotTableSourceType.xlScenario;
                    break;
            }
            return result;
        }
    }
}
