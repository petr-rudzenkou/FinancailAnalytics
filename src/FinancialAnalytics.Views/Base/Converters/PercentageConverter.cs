using System;
using System.Text;
using System.Windows.Data;

namespace FinancialAnalytics.Views.Base.Converters
{
    public class PercentageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            StringBuilder result = new StringBuilder();
            string stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                result.Append("(");
                result.Append(stringValue);
                result.Append("%");
                result.Append(")");
            }
            return result.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
