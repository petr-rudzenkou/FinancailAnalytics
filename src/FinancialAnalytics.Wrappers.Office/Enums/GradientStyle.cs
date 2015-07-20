using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{ 
    /// <summary>
    /// Specifies the style for a gradient fill.
    /// </summary>
    public enum GradientStyle
    {
        GradientDiagonalDown = 4, //Diagonal gradient moving from a top corner down to the opposite corner.
        GradientDiagonalUp = 3, //Diagonal gradient moving from a bottom corner up to the opposite corner.
        GradientFromCenter = 7, //Gradient running from the center out to the corners.
        GradientFromCorner = 5, //Gradient running from a corner to the other three corners.
        GradientFromTitle = 6, //Gradient running from the title outward.
        GradientHorizontal = 1, //Gradient running horizontally across the shape.
        GradientMixed = -2, //Gradient is mixed.
        GradientVertical = 2 //Gradient running vertically down the shape.
    }
}
