using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum SaveAsAccessMode
    {
        Exclusive = 3, //	Exclusive mode
        NoChange = 1, //	Default (does not change the access mode)
        Shared = 2, //	Share list
    }
}
