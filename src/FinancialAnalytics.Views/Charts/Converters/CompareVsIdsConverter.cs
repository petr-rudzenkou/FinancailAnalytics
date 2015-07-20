using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Data;

namespace FinancialAnalytics.Views.Charts.Converters
{
    public class CompareVsIdsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string result = string.Empty;
            var ids = value as List<string>;
            if (ids != null)
            {
                result = string.Join(", ", ids);
            }
            return result;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var result = new List<string>();
            var stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                var trimedSymbols = stringValue.Trim();
                if (string.IsNullOrEmpty(trimedSymbols))
                    return result;

                result = trimedSymbols.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
            }
            return result;
        }
    }
}
