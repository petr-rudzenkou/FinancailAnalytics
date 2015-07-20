using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum TextOrientation
    {
        /// <summary>
        /// Not supported.
        /// </summary>
        TextOrientationMixed,

        /// <summary>
        /// Horizontal.
        /// </summary>
        TextOrientationHorizontal,
        
        /// <summary>
        /// Upward.
        /// </summary>
        TextOrientationUpward,

        /// <summary>
        /// Downward.
        /// </summary>
        TextOrientationDownward,
        
        /// <summary>
        /// Vertical as required for Far East language support.
        /// </summary>
        TextOrientationVerticalFarEast,
        
        /// <summary>
        /// Vertical.
        /// </summary>
        TextOrientationVertical,
        
        /// <summary>
        /// Horizontal and rotated as required for Far East language support.
        /// </summary>
        TextOrientationHorizontalRotatedFarEast
    }
}
