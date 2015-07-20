using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Enums
{
    /// <summary>
    /// Excel analog is XLPivotCellType
    /// </summary>
    public enum PivotCellType
    {
        PivotCellBlankCell, //A structural blank cell in the PivotTable.

        PivotCellCustomSubtotal, //A cell in the row or column area that is a custom subtotal.

        PivotCellDataField, //A data field label (not the Data button).

        PivotCellDataPivotField, //The Data button.

        PivotCellGrandTotal, //A cell in a row or column area which is a grand total.

        PivotCellPageFieldItem, //The cell that shows the selected item of a Page field.

        PivotCellPivotField, //The button for a field (not the Data button).

        PivotCellPivotItem, //A cell in the row or column area which is not a subtotal, grand total, custom subtotal, or blank line.

        PivotCellSubtotal, //A cell in the row or column area which is a subtotal.

        PivotCellValue, //Any cell in the data area (except a blank row).
    }
}
