using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum PivotTableVersionList
    {
        Version2000 = 0, 	//Excel 2000
        Version10 = 1, 	//Excel 2002
        Version11 = 2, 	//Excel 2003
        Version12 = 3, 	//Excel 2007
	    TableVersion14 = 4, 	//Excel 2010
	    VersionCurrent = -1, 	//Provided only for backward compatibility
    }
}
