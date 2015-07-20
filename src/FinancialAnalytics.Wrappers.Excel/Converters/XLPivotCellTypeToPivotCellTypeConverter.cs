using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPivotCellTypeToPivotCellTypeConverter
    {
        public PivotCellType Convert(XlPivotCellType xlPivotCellType)
        {
            PivotCellType result;
            switch (xlPivotCellType)
            {
                case XlPivotCellType.xlPivotCellBlankCell:
                    result = PivotCellType.PivotCellBlankCell;
                    break;
                case XlPivotCellType.xlPivotCellCustomSubtotal:
                    result = PivotCellType.PivotCellCustomSubtotal;
                    break;
                case XlPivotCellType.xlPivotCellDataField:
                    result = PivotCellType.PivotCellDataField;
                    break;
                case XlPivotCellType.xlPivotCellDataPivotField:
                    result = PivotCellType.PivotCellDataPivotField;
                    break;
                case XlPivotCellType.xlPivotCellGrandTotal:
                    result = PivotCellType.PivotCellGrandTotal;
                    break;
                case XlPivotCellType.xlPivotCellPageFieldItem:
                    result = PivotCellType.PivotCellPageFieldItem;
                    break;
                case XlPivotCellType.xlPivotCellPivotField:
                    result = PivotCellType.PivotCellPivotField;
                    break;
                case XlPivotCellType.xlPivotCellPivotItem:
                    result = PivotCellType.PivotCellPivotItem;
                    break;
                case XlPivotCellType.xlPivotCellSubtotal:
                    result = PivotCellType.PivotCellSubtotal;
                    break;
                case XlPivotCellType.xlPivotCellValue:
                    result = PivotCellType.PivotCellValue;
                    break;
                default:
                    throw new InvalidEnumArgumentException("xlPivotCellType");
            }
            return result;
        }

        public XlPivotCellType ConvertBack(PivotCellType pivotCellType)
        {
            XlPivotCellType result;
            switch (pivotCellType)
            {
                case PivotCellType.PivotCellBlankCell:
                    result = XlPivotCellType.xlPivotCellBlankCell;
                    break;
                case PivotCellType.PivotCellCustomSubtotal:
                    result = XlPivotCellType.xlPivotCellCustomSubtotal;
                    break;
                case PivotCellType.PivotCellDataField:
                    result = XlPivotCellType.xlPivotCellDataField;
                    break;
                case PivotCellType.PivotCellDataPivotField:
                    result = XlPivotCellType.xlPivotCellDataPivotField;
                    break;
                case PivotCellType.PivotCellGrandTotal:
                    result = XlPivotCellType.xlPivotCellGrandTotal;
                    break;
                case PivotCellType.PivotCellPageFieldItem:
                    result = XlPivotCellType.xlPivotCellPageFieldItem;
                    break;
                case PivotCellType.PivotCellPivotField:
                    result = XlPivotCellType.xlPivotCellPivotField;
                    break;
                case PivotCellType.PivotCellPivotItem:
                    result = XlPivotCellType.xlPivotCellPivotItem;
                    break;
                case PivotCellType.PivotCellSubtotal:
                    result = XlPivotCellType.xlPivotCellSubtotal;
                    break;
                case PivotCellType.PivotCellValue:
                    result = XlPivotCellType.xlPivotCellValue;
                    break;
                default:
                    throw new InvalidEnumArgumentException("pivotCellType");
            }
            return result;

        }
    }
}
