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
  public class TickLabelPositionToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
          if (value is TickLabelPosition)
            {

                var tickLabelPositionValue = (TickLabelPosition)(value);
              
              switch (tickLabelPositionValue)
                {
                    case TickLabelPosition.None:
                        return Content.TickLabelPosition_None;
                    case TickLabelPosition.Low:
                        return Content.TickLabelPosition_Low;
                    case TickLabelPosition.High:
                        return Content.TickLabelPosition_High;
                    case TickLabelPosition.NextToAxis:
                        return Content.TickLabelPosition_NextToAxis;
                    default:
                        return Content.TickLabelPosition_None;
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
