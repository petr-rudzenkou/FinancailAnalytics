using System;
using System.Windows.Data;
using System.Windows.Media;

namespace FinancialAnalytics.Views.Base.Converters
{
    public class PriceGainFontColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            SolidColorBrush result = new SolidColorBrush(Colors.Black);
            string stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                if (stringValue.StartsWith("+"))
                {
                    result = new SolidColorBrush(Colors.Green);
                }
                else
                {
                    result = new SolidColorBrush(Colors.Red);
                }
            }
            return result;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
