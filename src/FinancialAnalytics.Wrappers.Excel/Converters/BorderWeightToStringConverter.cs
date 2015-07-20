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
    public class BorderWeightToStringConverter:IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is BorderWeight)
            {

                var borderWeightValue = (BorderWeight) (value);

                switch (borderWeightValue)
                {
                    case BorderWeight.None:
                        return Content.BorderWeight_None;
                    case BorderWeight.Hairline:
                        return Content.BorderWeight_Hairline;
                    case BorderWeight.Medium:
                        return Content.BorderWeight_Medium;
                    case BorderWeight.Thick:
                        return Content.BorderWeight_Thick;
                    case BorderWeight.Thin:
                        return Content.BorderWeight_Thin;
                    default:
                        return Content.BorderWeight_None;
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
