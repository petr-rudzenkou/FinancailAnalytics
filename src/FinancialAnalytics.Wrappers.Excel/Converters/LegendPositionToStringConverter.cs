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
   public class LegendPositionToStringConverter:IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
           if (value is LegendPosition)
            {

                var legendPositionValue = (LegendPosition)(value);
               
               switch (legendPositionValue)
                {
                    case LegendPosition.LegendPositionBottom:
                        return Content.LegendPosition_LegendPositionBottom;
                    case LegendPosition.LegendPositionCorner:
                        return Content.LegendPosition_LegendPositionCorner;
                    case LegendPosition.LegendPositionCustom:
                        return Content.LegendPosition_LegendPositionCustom;
                    case LegendPosition.LegendPositionLeft:
                        return Content.LegendPosition_LegendPositionLeft;
                    case LegendPosition.LegendPositionRight:
                        return Content.LegendPosition_LegendPositionRight;
                    case LegendPosition.LegendPositionTop:
                        return Content.LegendPosition_LegendPositionTop;
                    default:
                        return Content.LegendPosition_LegendPositionBottom;
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
