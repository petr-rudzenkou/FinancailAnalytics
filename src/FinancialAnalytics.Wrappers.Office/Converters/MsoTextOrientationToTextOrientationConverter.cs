using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoTextOrientationToTextOrientationConverter
    {
        public static TextOrientation Convert(MsoTextOrientation msoTextOrientation)
        {
            TextOrientation textOrientation;
            switch (msoTextOrientation)
            {
                case MsoTextOrientation.msoTextOrientationDownward:
                    textOrientation = TextOrientation.TextOrientationDownward;
                    break;
                case MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast:
                    textOrientation = TextOrientation.TextOrientationHorizontalRotatedFarEast;
                    break;
                case MsoTextOrientation.msoTextOrientationMixed:
                    textOrientation = TextOrientation.TextOrientationMixed;
                    break;
                case MsoTextOrientation.msoTextOrientationUpward:
                    textOrientation = TextOrientation.TextOrientationUpward;
                    break;
                case MsoTextOrientation.msoTextOrientationVertical:
                    textOrientation = TextOrientation.TextOrientationVertical;
                    break;
                case MsoTextOrientation.msoTextOrientationVerticalFarEast:
                    textOrientation = TextOrientation.TextOrientationVerticalFarEast;
                    break;
                default:
                    textOrientation = TextOrientation.TextOrientationHorizontal;
                    break;
            }
            return textOrientation;
        }

        public static MsoTextOrientation ConvertBack(TextOrientation textOrientation)
        {
            MsoTextOrientation msoTextOrientation;
            switch (textOrientation)
            {
                case TextOrientation.TextOrientationDownward:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationDownward;
                    break;
                case TextOrientation.TextOrientationHorizontalRotatedFarEast:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast;
                    break;
                case TextOrientation.TextOrientationMixed:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationMixed;
                    break;
                case TextOrientation.TextOrientationUpward:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationUpward;
                    break;
                case TextOrientation.TextOrientationVertical:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationVertical;
                    break;
                case TextOrientation.TextOrientationVerticalFarEast:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationVerticalFarEast;
                    break;
                default:
                    msoTextOrientation = MsoTextOrientation.msoTextOrientationHorizontal;
                    break;
            }
            return msoTextOrientation;
        }
    }
}
