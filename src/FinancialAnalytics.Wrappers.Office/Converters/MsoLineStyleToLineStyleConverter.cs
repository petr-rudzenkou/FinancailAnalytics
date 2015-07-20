using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoLineStyleToLineStyleConverter
    {
        public static LineStyle Convert(MsoLineStyle msoLineStyle)
        {
            LineStyle lineStyle;
            switch (msoLineStyle)
            {
                case MsoLineStyle.msoLineStyleMixed:
                    lineStyle = LineStyle.LineStyleMixed;
                    break;
                case MsoLineStyle.msoLineThickBetweenThin:
                    lineStyle = LineStyle.LineThickBetweenThin;
                    break;
                case MsoLineStyle.msoLineThickThin:
                    lineStyle = LineStyle.LineThickThin;
                    break;
                case MsoLineStyle.msoLineThinThick:
                    lineStyle = LineStyle.LineThinThick;
                    break;
                case MsoLineStyle.msoLineThinThin:
                    lineStyle = LineStyle.LineThinThin;
                    break;
                default:
                    lineStyle = LineStyle.LineSingle;
                    break;
            }
            return lineStyle;
        }

        public static MsoLineStyle ConvertBack(LineStyle lineStyle)
        {
            MsoLineStyle msoLineStyle;
            switch (lineStyle)
            {
                case LineStyle.LineStyleMixed:
                    msoLineStyle = MsoLineStyle.msoLineStyleMixed;
                    break;
                case LineStyle.LineThickBetweenThin:
                    msoLineStyle = MsoLineStyle.msoLineThickBetweenThin;
                    break;
                case LineStyle.LineThickThin:
                    msoLineStyle = MsoLineStyle.msoLineThickThin;
                    break;
                case LineStyle.LineThinThick:
                    msoLineStyle = MsoLineStyle.msoLineThinThick;
                    break;
                case LineStyle.LineThinThin:
                    msoLineStyle = MsoLineStyle.msoLineThinThin;
                    break;
                default:
                    msoLineStyle = MsoLineStyle.msoLineSingle;
                    break;
            }
            return msoLineStyle;
        }
    }
}
