using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    public enum PasteType
    {
        /// <summary>
        /// Only the values will be pasted.
        /// </summary>
        PasteValues,
        
        /// <summary>
        /// Comments will be pasted.
        /// </summary>
        PasteComments,
        
        /// <summary>
        /// Formulas will be pasted.
        /// </summary>
        PasteFormulas,
        
        /// <summary>
        /// Formatting will be pasted.
        /// </summary>
        PasteFormats,
        //
        // Summary:
        //     Everything will be pasted.
        PasteAll,
        
		/// <summary>
		/// Everything will be pasted with styles from source theme.
		/// </summary>
		PasteAllUsingSourceTheme,

        /// <summary>
        /// Validation from the source cell is applied to the destination cell.
        /// </summary>
        PasteValidation,
        
        /// <summary>
        /// Everything except borders will be pasted.
        /// </summary>
        PasteAllExceptBorders,
        
        /// <summary>
        /// The column width of the source cell will be applied to the destination cell.
        /// </summary>
        PasteColumnWidths,
        
        /// <summary>
        /// Formulas and number formats are pasted.
        /// </summary>
        PasteFormulasAndNumberFormats,
        
        /// <summary>
        /// Only the values number formats will be pasted.
        /// </summary>
        PasteValuesAndNumberFormats
    }
}
