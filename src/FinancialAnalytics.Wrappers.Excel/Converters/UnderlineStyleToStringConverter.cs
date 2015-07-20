using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Resources;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class UnderlineStyleToStringConverter:IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is UnderlineStyle)
            {

                var underlineStyleValue = (UnderlineStyle)(value);

                switch (underlineStyleValue)
                {
                    case UnderlineStyle.UnderlineStyleNone:
                        return Content.UnderlineStyle_UnderlineStyleNone;
                    case UnderlineStyle.UnderlineStyleDouble:
                        return Content.UnderlineStyle_UnderlineStyleDouble;
                    case UnderlineStyle.UnderlineStyleDoubleAccounting:
                        return Content.UnderlineStyle_UnderlineStyleDoubleAccounting;
                    case UnderlineStyle.UnderlineStyleSingle:
                        return Content.UnderlineStyle_UnderlineStyleSingle;
                    case UnderlineStyle.UnderlineStyleSingleAccounting:
                        return Content.UnderlineStyle_UnderlineStyleSingleAccounting;
                    default:
                        return Content.UnderlineStyle_UnderlineStyleNone;
                }
            }
            throw new ArgumentException();
           

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
