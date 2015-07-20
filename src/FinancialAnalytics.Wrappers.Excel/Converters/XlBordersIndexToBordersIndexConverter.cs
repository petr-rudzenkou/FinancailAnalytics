using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlBordersIndexToBordersIndexConverter
    {
        public static BordersIndex Convert(XlBordersIndex xlBordersIndex)
        {
            BordersIndex bordersIndex;
            switch (xlBordersIndex)
            {
                case XlBordersIndex.xlDiagonalDown:
                    bordersIndex = BordersIndex.DiagonalDown;
                    break;
                case XlBordersIndex.xlDiagonalUp:
                    bordersIndex = BordersIndex.DiagonalUp;
                    break;
                case XlBordersIndex.xlEdgeBottom:
                    bordersIndex = BordersIndex.EdgeBottom;
                    break;
                case XlBordersIndex.xlEdgeLeft:
                    bordersIndex = BordersIndex.EdgeLeft;
                    break;
                case XlBordersIndex.xlEdgeRight:
                    bordersIndex = BordersIndex.EdgeRight;
                    break;
                case XlBordersIndex.xlEdgeTop:
                    bordersIndex = BordersIndex.EdgeTop;
                    break;
                case XlBordersIndex.xlInsideHorizontal:
                    bordersIndex = BordersIndex.InsideHorizontal;
                    break;
                default:
                    bordersIndex = BordersIndex.InsideVertical;
                    break;
            }
            return bordersIndex;
        }

        public static XlBordersIndex ConvertBack(BordersIndex bordersIndex)
        {
            XlBordersIndex xlBordersIndex;
            switch (bordersIndex)
            {
                case BordersIndex.DiagonalDown:
                    xlBordersIndex = XlBordersIndex.xlDiagonalDown;
                    break;
                case BordersIndex.DiagonalUp:
                    xlBordersIndex = XlBordersIndex.xlDiagonalUp;
                    break;
                case BordersIndex.EdgeBottom:
                    xlBordersIndex = XlBordersIndex.xlEdgeBottom;
                    break;
                case BordersIndex.EdgeLeft:
                    xlBordersIndex = XlBordersIndex.xlEdgeLeft;
                    break;
                case BordersIndex.EdgeRight:
                    xlBordersIndex = XlBordersIndex.xlEdgeRight;
                    break;
                case BordersIndex.EdgeTop:
                    xlBordersIndex = XlBordersIndex.xlEdgeTop;
                    break;
                case BordersIndex.InsideHorizontal:
                    xlBordersIndex = XlBordersIndex.xlInsideHorizontal;
                    break;
                default:
                    xlBordersIndex = XlBordersIndex.xlInsideVertical;
                    break;
            }
            return xlBordersIndex;
        }
    }
}
