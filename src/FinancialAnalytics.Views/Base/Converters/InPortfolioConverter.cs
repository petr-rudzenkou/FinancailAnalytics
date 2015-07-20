using System;
using System.Linq;
using System.Windows.Data;
using FinancialAnalytics.Utils;

namespace FinancialAnalytics.Views.Base.Converters
{
    public class InPortfolioConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var result = true;
            var stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                result =  !PortfolioCacheProvider.PortfolioSymbols.Contains(stringValue);
            }
            return result;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
