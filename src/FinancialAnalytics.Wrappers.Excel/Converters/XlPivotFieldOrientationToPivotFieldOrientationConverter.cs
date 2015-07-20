using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPivotFieldOrientationToPivotFieldOrientationConverter
    {
        public static PivotFieldOrientation Convert(XlPivotFieldOrientation xlPivotFieldOrientation)
        {
            PivotFieldOrientation result;
            switch (xlPivotFieldOrientation)
            {
                case XlPivotFieldOrientation.xlColumnField:
                    result = PivotFieldOrientation.ColumnField;
                    break;
                case XlPivotFieldOrientation.xlDataField:
                    result = PivotFieldOrientation.DataField;
                    break;
                case XlPivotFieldOrientation.xlHidden:
                    result = PivotFieldOrientation.Hidden;
                    break;
                case XlPivotFieldOrientation.xlPageField:
                    result = PivotFieldOrientation.PageField;
                    break;
                case XlPivotFieldOrientation.xlRowField:
                    result = PivotFieldOrientation.RowField;
                    break;
                default:
                    throw new InvalidEnumArgumentException("xlPivotFieldOrientation");
            }
            return result;
        }

        public static XlPivotFieldOrientation ConvertBack(PivotFieldOrientation pivotFieldOrientation)
        {
            XlPivotFieldOrientation result;
            switch (pivotFieldOrientation)
            {
                case PivotFieldOrientation.ColumnField:
                    result = XlPivotFieldOrientation.xlColumnField;
                    break;
                case PivotFieldOrientation.DataField:
                    result = XlPivotFieldOrientation.xlDataField;
                    break;
                case PivotFieldOrientation.Hidden:
                    result = XlPivotFieldOrientation.xlHidden;
                    break;
                case PivotFieldOrientation.PageField:
                    result = XlPivotFieldOrientation.xlPageField;
                    break;
                case PivotFieldOrientation.RowField:
                    result = XlPivotFieldOrientation.xlRowField;
                    break;
                default:
                    throw new InvalidEnumArgumentException("pivotFieldOrientation");
            }
            return result;
        }
    }
}
