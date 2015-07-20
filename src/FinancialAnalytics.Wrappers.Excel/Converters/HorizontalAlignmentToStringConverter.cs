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
   public class HorizontalAlignmentToStringConverter:IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
           if (value is HorizontalAlignment)
            {

                var horizontalAlignmentValue = (HorizontalAlignment)(value);

                switch (horizontalAlignmentValue)
                {
                    case HorizontalAlignment.Right:
                        return Content.HorizontalAlignment_Right;
                    case HorizontalAlignment.Left:
                        return Content.HorizontalAlignment_Left;
                    case HorizontalAlignment.Justify:
                        return Content.HorizontalAlignment_Justify;
                    case HorizontalAlignment.Distributed:
                        return Content.HorizontalAlignment_Distributed;
                    case HorizontalAlignment.Center:
                        return Content.HorizontalAlignment_Center;
                    case HorizontalAlignment.CenterAcrossSelection:
                        return Content.HorizontalAlignment_CenterAcrossSelection;
                    case HorizontalAlignment.Fill:
                        return Content.HorizontalAlignment_Fill;
                    case HorizontalAlignment.General:
                        return Content.HorizontalAlignment_General;
                    default:
                        return Content.HorizontalAlignment_Right;
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
