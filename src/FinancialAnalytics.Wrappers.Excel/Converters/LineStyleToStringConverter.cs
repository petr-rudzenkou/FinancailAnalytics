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
   public class LineStyleConverter:IValueConverter
   {
       public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
       {
           if (value is LineStyle)
           {

               var lineStyleValue = (LineStyle) (value);

               switch (lineStyleValue)
               {
                   case LineStyle.LineStyleNone:
                       return Content.LineStyle_LineStyleNone;
                   case LineStyle.Continuous:
                       return Content.LineStyle_Continuous;
                   case LineStyle.Dash:
                       return Content.LineStyle_Dash;
                   case LineStyle.DashDot:
                       return Content.LineStyle_DashDot;
                   case LineStyle.DashDotDot:
                       return Content.LineStyle_DashDotDot;
                   case LineStyle.Dot:
                       return Content.LineStyle_Dot;
                   case LineStyle.Double:
                       return Content.LineStyle_Double;
                   case LineStyle.SlantDashDot:
                       return Content.LineStyle_SlantDashDot;
                   default:
                       return Content.LineStyle_LineStyleNone;
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
