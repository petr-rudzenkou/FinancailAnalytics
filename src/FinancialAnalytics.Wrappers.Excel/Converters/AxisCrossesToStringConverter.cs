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
   public class AxisCrossesToStringConverter:IValueConverter
   {
       public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
       {
           if (value is AxisCrosses)
           {
               var axisCrossesValue = (AxisCrosses)(value);

               switch (axisCrossesValue)
               {
                   case AxisCrosses.AxisCrossesAutomatic:
                       return Content.AxisCrosses_Automatic;
                   case AxisCrosses.AxisCrossesCustom:
                       return Content.AxisCrosses_Custom;
                   case AxisCrosses.AxisCrossesMaximum:
                       return Content.AxisCrosses_Maximum;
                   case AxisCrosses.AxisCrossesMinimum:
                       return Content.AxisCrosses_Minimum;
                   default:
                       return Content.AxisCrosses_Automatic;
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
