using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlLegendPositionToLegendPositionConverter
    {

        public static LegendPosition Convert(int xlLegendPosition)
        {
            LegendPosition legendPosition;
            switch (xlLegendPosition)
            {
                case -4107 : //XlLegendPosition.xlLegendPositionBottom
                    legendPosition = LegendPosition.LegendPositionBottom;
                    break;
                case 2 : //XlLegendPosition.xlLegendPositionCorner
                    legendPosition = LegendPosition.LegendPositionCorner;
                    break;
                case -4131 : // XlLegendPosition.xlLegendPositionLeft
                    legendPosition = LegendPosition.LegendPositionLeft;
                    break;
                case -4152 : //XlLegendPosition.xlLegendPositionRight
                    legendPosition = LegendPosition.LegendPositionRight;
                    break;
                case -4161: //XlLegendPosition.xlLegendPositionCustom
                    legendPosition = LegendPosition.LegendPositionCustom;
                    break;
                case -4160 : //XlLegendPosition.xlLegendPositionTop
                    legendPosition = LegendPosition.LegendPositionTop;
                    break;

                default:
                    legendPosition = LegendPosition.LegendPositionCustom;
                    break;
            }
            return legendPosition;
        }

        public static int ConvertBack(LegendPosition legendPosition)
        {
            int xlLegendPosition;
            switch (legendPosition)
            {
                case LegendPosition.LegendPositionBottom :
                    xlLegendPosition = -4107; //XlLegendPosition.xlLegendPositionBottom
                    break;
                case LegendPosition.LegendPositionCorner :
                    xlLegendPosition = 2; //XlLegendPosition.xlLegendPositionCorner
                    break;
                case LegendPosition.LegendPositionLeft :
                    xlLegendPosition = -4131; // XlLegendPosition.xlLegendPositionLeft
                    break;
                case LegendPosition.LegendPositionRight :
                    xlLegendPosition = -4152; //XlLegendPosition.xlLegendPositionRight
                    break;
                case LegendPosition.LegendPositionTop :
                    xlLegendPosition = -4160; //XlLegendPosition.xlLegendPositionTop
                    break;
                case LegendPosition.LegendPositionCustom :
                    xlLegendPosition =  -4161; //XlLegendPosition.xlLegendPositionCustom
                    break;
                default:
                    xlLegendPosition = -4161; //XlLegendPosition.xlLegendPositionCustom
                    break;
            }
            return xlLegendPosition;
        }
    }
}
