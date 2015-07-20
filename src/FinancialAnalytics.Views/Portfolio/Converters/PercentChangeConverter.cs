using System;
using System.Windows.Data;

namespace FinancialAnalytics.Views.Portfolio.Converters
{
    public class PercentChangeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string str = value as string;
            if (!string.IsNullOrEmpty(str))
                return str + "%";
            return str;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
