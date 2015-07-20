using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum CmdType
    {
        Cube, 	//Contains a cube name for an OLAP data source.
        Default, 	//Contains command text that the OLE DB provider understands
        List, 	//Contains a pointer to list data.
        Sql, 	//Contains an SQL statement.
        Table, 	//Contains a table name for accessing OLE DB data sources.
    }
}
