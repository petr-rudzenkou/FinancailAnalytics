using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum ScaleFrom
    {
        /// <summary>
        /// Shape's top left corner retains its position.
        /// </summary>
        ScaleFromTopLeft,
        
        /// <summary>
        /// Shape's midpoint retains its position.
        /// </summary>
        ScaleFromMiddle,
        
        /// <summary>
        /// Shape's bottom right corner retains its position.
        /// </summary>
        ScaleFromBottomRight
    }
}
