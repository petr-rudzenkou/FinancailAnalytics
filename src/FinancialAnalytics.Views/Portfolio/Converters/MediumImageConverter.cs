using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace FinancialAnalytics.Views.Portfolio.Converters
{
    public class MediumImageConverter : IValueConverter
    {
        private const string IMAGE_URL_FORMAT_STRING = @"https://chart.finance.yahoo.com/c/1y/{0}?lang=en-US&region=US";

        public object Convert(object value, Type targetType,
                          object parameter, CultureInfo culture)
        {
            string symbol = value as string;
            if (symbol != null)
            {
                var uri = new Uri(string.Format(IMAGE_URL_FORMAT_STRING, symbol));
                var image = new BitmapImage(uri);
                return image;
            }
            return new BitmapImage();
        }

        public object ConvertBack(object value, Type targetType,
                                  object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
