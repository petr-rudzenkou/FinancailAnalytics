using System;
using System.Text;
using System.Windows.Data;

namespace FinancialAnalytics.Views.Base.Converters
{
    public class PriceGainConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            StringBuilder result = new StringBuilder();
            string stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                string[] priceChanges = stringValue.Split(new[] {" - "}, StringSplitOptions.RemoveEmptyEntries);
                if (priceChanges.Length == 2)
                {
                    result.Append(priceChanges[0]);
                    result.Append("(");
                    result.Append(priceChanges[1]);
                    result.Append(")");
                }
            }
            return result.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
