using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    class XlDataLabelPositionToDataLabelPosition
    {
        public static DataLabelPosition Convert(XlDataLabelPosition xlDataLabelPosition)
        {
            DataLabelPosition dataLabelPosition;
            switch (xlDataLabelPosition)
            {
                case XlDataLabelPosition.xlLabelPositionAbove:
                    dataLabelPosition = DataLabelPosition.LabelPositionAbove;
                    break;
                case XlDataLabelPosition.xlLabelPositionBelow:
                    dataLabelPosition = DataLabelPosition.LabelPositionBelow;
                    break;
                case XlDataLabelPosition.xlLabelPositionBestFit:
                    dataLabelPosition = DataLabelPosition.LabelPositionBestFit;
                    break;
                case XlDataLabelPosition.xlLabelPositionCenter:
                    dataLabelPosition = DataLabelPosition.LabelPositionCenter;
                    break;
                case XlDataLabelPosition.xlLabelPositionCustom:
                    dataLabelPosition = DataLabelPosition.LabelPositionCustom;
                    break;
                case XlDataLabelPosition.xlLabelPositionInsideBase:
                    dataLabelPosition = DataLabelPosition.LabelPositionInsideBase;
                    break;
                case XlDataLabelPosition.xlLabelPositionInsideEnd:
                    dataLabelPosition = DataLabelPosition.LabelPositionInsideEnd;
                    break;
                case XlDataLabelPosition.xlLabelPositionLeft:
                    dataLabelPosition = DataLabelPosition.LabelPositionLeft;
                    break;
                case XlDataLabelPosition.xlLabelPositionMixed:
                    dataLabelPosition = DataLabelPosition.LabelPositionMixed;
                    break;
                case XlDataLabelPosition.xlLabelPositionOutsideEnd:
                    dataLabelPosition = DataLabelPosition.LabelPositionOutsideEnd;
                    break;
                case XlDataLabelPosition.xlLabelPositionRight:
                    dataLabelPosition = DataLabelPosition.LabelPositionRight;
                    break;
                default:
                    dataLabelPosition = DataLabelPosition.LabelPositionAbove;
                    break;
            }
            return dataLabelPosition;
        }

        public static XlDataLabelPosition ConvertBack(DataLabelPosition dataLabelPosition)
        {
            XlDataLabelPosition xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionAbove;
            switch (dataLabelPosition)
            {
                case DataLabelPosition.LabelPositionAbove:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionAbove;
                    break;
                case DataLabelPosition.LabelPositionBelow:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionBelow;
                    break;
                case DataLabelPosition.LabelPositionBestFit:
                    xlDataLabelPosition =XlDataLabelPosition.xlLabelPositionBestFit;
                    break;
                case DataLabelPosition.LabelPositionCenter:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionCenter;
                    break;
                case DataLabelPosition.LabelPositionCustom:
                    dataLabelPosition = DataLabelPosition.LabelPositionCustom;
                    break;
                case DataLabelPosition.LabelPositionInsideBase:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionInsideBase;
                    break;
                case DataLabelPosition.LabelPositionInsideEnd:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionInsideEnd;
                    break;
                case DataLabelPosition.LabelPositionLeft:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionLeft;
                    break;
                case DataLabelPosition.LabelPositionMixed:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionMixed;
                    break;
                case DataLabelPosition.LabelPositionOutsideEnd:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionOutsideEnd;
                    break;
                case DataLabelPosition.LabelPositionRight:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionRight;
                    break;
                default:
                    xlDataLabelPosition = XlDataLabelPosition.xlLabelPositionAbove;
                    break;
            }

            return xlDataLabelPosition;
        }
    }
}
