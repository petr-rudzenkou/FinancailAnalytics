using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum PivotTableSourceType
    {
        Consolidation =	3, //	Multiple consolidation ranges.
        Database =	1, //	Microsoft Excel list or database.
        External = 2, 	//Data from another application.
        PivotTable = -4148, //	Same source as another PivotTable report.
        Scenario = 4, //	Data is based on scenarios created using the Scenario Manager.
    }
}
