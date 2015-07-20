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
   public class VerticalAlignmentConverter:IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            if (value is VerticalAlignment)
            {

                var verticalAlignmentValue = (VerticalAlignment)(value);

                switch (verticalAlignmentValue)
                {
                    case VerticalAlignment.Bottom:
                        return Content.VerticalAlignment_Bottom;
                    case VerticalAlignment.Center:
                        return Content.VerticalAlignment_Center;
                    case VerticalAlignment.Distributed:
                        return Content.VerticalAlignment_Distributed;
                    case VerticalAlignment.Justify:
                        return Content.VerticalAlignment_Justify;
                    case VerticalAlignment.Top:
                        return Content.VerticalAlignment_Top;
                    default:
                        return Content.VerticalAlignment_Bottom;
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
