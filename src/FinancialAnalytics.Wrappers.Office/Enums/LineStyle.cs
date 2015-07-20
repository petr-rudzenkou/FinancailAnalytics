using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum LineStyle
    {
        /// <summary>
        /// Not supported.
        /// </summary>
        LineStyleMixed,
        
        /// <summary>
        /// Single line.
        /// </summary>
        LineSingle,
        
        /// <summary>
        /// Two thin lines.
        /// </summary>
        LineThinThin,
        
        /// <summary>
        /// Thick line next to thin line. For horizontal lines, thick line is below thin
        /// line. For vertical lines, thick line is to the right of the thin line.
        /// </summary>
        LineThinThick,
        
        /// <summary>
        /// Thick line next to thin line. For horizontal lines, thick line is above thin
        /// line. For vertical lines, thick line is to the left of the thin line.
        /// </summary>
        LineThickThin,
        
        /// <summary>
        /// Thick line with a thin line on each side.
        /// </summary>
        LineThickBetweenThin
    }
}
