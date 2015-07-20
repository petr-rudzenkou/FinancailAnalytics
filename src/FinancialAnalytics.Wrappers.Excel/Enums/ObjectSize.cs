using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum ObjectSize
    {
        /// <summary>
        /// Print the chart the same size as it appears on the screen.
        /// </summary>
        ScreenSize,
        
        /// <summary>
        /// Print the chart as large as possible, while retaining the chart's height-to-width ratio as shown on the screen
        /// </summary>
        FitToPage,
        
        /// <summary>
        /// Print the chart to fit the page, adjusting the height-to-width ratio as necessary.
        /// </summary>
        FullPage
    }
}
