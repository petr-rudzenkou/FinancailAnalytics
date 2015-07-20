using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum PasteSpecialOperation
    {
        /// <summary>
        /// No calculation will be done in the paste operation.
        /// </summary>
        PasteSpecialOperationNone,
        
        /// <summary>
        /// Copied data will be added with the value in the destination cell.
        /// </summary>
        PasteSpecialOperationAdd,
        
        /// <summary>
        /// Copied data will be subtracted with the value in the destination cell.
        /// </summary>
        PasteSpecialOperationSubtract,
        
        /// <summary>
        /// Copied data will be multiplied with the value in the destination cell.
        /// </summary>
        PasteSpecialOperationMultiply,
        
        /// <summary>
        /// Copied data will be divided with the value in the destination cell.
        /// </summary>
        PasteSpecialOperationDivide
    }
}
