using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlUnderlineStyleToUnderlineStyleConverter
    {

        public static UnderlineStyle Convert(XlUnderlineStyle xlUnderlineStyle)
        {
            UnderlineStyle underlineStyle;
            switch (xlUnderlineStyle)
            {
                case XlUnderlineStyle.xlUnderlineStyleDouble:
                    underlineStyle = UnderlineStyle.UnderlineStyleDouble;
                    break;
                case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
                    underlineStyle = UnderlineStyle.UnderlineStyleDoubleAccounting;
                    break;
                case XlUnderlineStyle.xlUnderlineStyleNone:
                    underlineStyle = UnderlineStyle.UnderlineStyleNone;
                    break;
                case XlUnderlineStyle.xlUnderlineStyleSingle:
                    underlineStyle = UnderlineStyle.UnderlineStyleSingle;
                    break;
                case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
                    underlineStyle = UnderlineStyle.UnderlineStyleSingleAccounting;
                    break;
                default:
                    underlineStyle = UnderlineStyle.UnderlineStyleNone;
                    break;
            }
            return underlineStyle;
        }

        public static XlUnderlineStyle ConvertBack(UnderlineStyle underlineStyle)
        {
            XlUnderlineStyle xlUnderlineStyle;
            switch (underlineStyle)
            {
                case UnderlineStyle.UnderlineStyleDouble:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleDouble;
                    break;
                case UnderlineStyle.UnderlineStyleDoubleAccounting:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleDoubleAccounting;
                    break;
                case UnderlineStyle.UnderlineStyleNone:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleNone;
                    break;
                case UnderlineStyle.UnderlineStyleSingle:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleSingle;
                    break;
                case UnderlineStyle.UnderlineStyleSingleAccounting:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleSingleAccounting;
                    break;
                default:
                    xlUnderlineStyle = XlUnderlineStyle.xlUnderlineStyleNone;
                    break;
            }
            return xlUnderlineStyle;
        }
    }
}
