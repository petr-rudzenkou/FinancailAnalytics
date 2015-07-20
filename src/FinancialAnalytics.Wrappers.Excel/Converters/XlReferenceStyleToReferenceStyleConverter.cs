using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public static class XlReferenceStyleToReferenceStyleConverter
    {
        public static ReferenceStyle Convert(XlReferenceStyle xlStyle)
        {
            return (ReferenceStyle) (Int32) xlStyle;
        }

        public static XlReferenceStyle ConvertBack(ReferenceStyle style)
        {
            return (XlReferenceStyle) (Int32) style;
        }
    }
}
