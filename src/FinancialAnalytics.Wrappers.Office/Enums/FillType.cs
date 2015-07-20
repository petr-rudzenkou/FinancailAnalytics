using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum FillType
    {
        /// <summary>
        /// Mixed fill.
        /// </summary>
        FillMixed,
        
        /// <summary>
        /// Solid fill.
        /// </summary>
        FillSolid,
        
        /// <summary>
        /// Patterned fill.
        /// </summary>
        FillPatterned,
        
        /// <summary>
        /// Gradient fill.
        /// </summary>
        FillGradient,
        
        /// <summary>
        /// Textured fill.
        /// </summary>
        FillTextured,
        
        /// <summary>
        /// Fill is the same as the background.
        /// </summary>
        FillBackground,
        
        /// <summary>
        /// Picture fill.
        /// </summary>
        FillPicture
    }
}
