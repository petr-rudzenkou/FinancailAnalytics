using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum PictureColorType
    {
        /// <summary>
        /// Mixed transformation.
        /// </summary>
        PictureMixed,
        
        /// <summary>
        /// Default color transformation.
        /// </summary>
        PictureAutomatic,
        
        /// <summary>
        /// Grayscale transformation.
        /// </summary>
        PictureGrayscale,
        
        /// <summary>
        /// Black-and-white transformation.
        /// </summary>
        PictureBlackAndWhite,
        
        /// <summary>
        /// Watermark transformation.
        /// </summary>
        PictureWatermark
    }
}
