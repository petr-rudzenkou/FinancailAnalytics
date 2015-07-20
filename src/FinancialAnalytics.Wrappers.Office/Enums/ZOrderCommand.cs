using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Office.Enums
{
    public enum ZOrderCommand
    {
        /// <summary>
        /// Bring shape to the front.
        /// </summary>
        BringToFront,
        
        /// <summary>
        /// Send shape to the back.
        /// </summary>
        SendToBack,
        
        /// <summary>
        /// Bring shape forward.
        /// </summary>
        BringForward,
        
        /// <summary>
        /// Send shape backward.
        /// </summary>
        SendBackward,
        
        /// <summary>
        /// Bring shape in front of text.
        /// </summary>
        BringInFrontOfText,
        
        /// <summary>
        /// Send shape behind text.
        /// </summary>
        SendBehindText
    }
}
