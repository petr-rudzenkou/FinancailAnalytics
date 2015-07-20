using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum ChartLocation
    {
        /// <summary>
        /// Chart is moved to a new sheet.
        /// </summary>
        LocationAsNewSheet,

        /// <summary>
        /// Chart is to be embedded in an existing sheet.
        /// </summary>
        LocationAsObject,
        
        /// <summary>
        /// Excel controls chart location.
        /// </summary>
        LocationAutomatic
    }
}
