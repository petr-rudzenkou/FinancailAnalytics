using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlObjectSizeToObjectSizeConverter
    {
        public static ObjectSize Convert(XlObjectSize xlObjectSize)
        {
            ObjectSize objectSize;
            switch (xlObjectSize)
            {
                case XlObjectSize.xlFitToPage:
                    objectSize = ObjectSize.FitToPage;
                    break;
                case XlObjectSize.xlFullPage:
                    objectSize = ObjectSize.FullPage;
                    break;
                default:
                    objectSize = ObjectSize.ScreenSize;
                    break;
            }
            return objectSize;
        }

        public static XlObjectSize ConvertBack(ObjectSize objectSize)
        {
            XlObjectSize xlObjectSize;
            switch (objectSize)
            {
                case ObjectSize.FitToPage:
                    xlObjectSize = XlObjectSize.xlFitToPage;
                    break;
                case ObjectSize.FullPage:
                    xlObjectSize = XlObjectSize.xlFullPage;
                    break;
                default:
                    xlObjectSize = XlObjectSize.xlScreenSize;
                    break;
            }
            return xlObjectSize;
        }
    }
}
