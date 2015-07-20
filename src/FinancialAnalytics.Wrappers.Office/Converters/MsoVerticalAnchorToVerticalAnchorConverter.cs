using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoVerticalAnchorToVerticalAnchorConverter
    {
        public static VerticalAnchor Convert(MsoVerticalAnchor msoVerticalAnchor)
        {
            VerticalAnchor verticalAnchor;
            switch (msoVerticalAnchor)
            {
                case MsoVerticalAnchor.msoAnchorBottom:
                    verticalAnchor = VerticalAnchor.AnchorBottom;
                    break;
                case MsoVerticalAnchor.msoAnchorBottomBaseLine:
                    verticalAnchor = VerticalAnchor.AnchorBottomBaseLine;
                    break;
                case MsoVerticalAnchor.msoAnchorMiddle:
                    verticalAnchor = VerticalAnchor.AnchorMiddle;
                    break;
                case MsoVerticalAnchor.msoAnchorTopBaseline:
                    verticalAnchor = VerticalAnchor.AnchorTopBaseline;
                    break;
                case MsoVerticalAnchor.msoVerticalAnchorMixed:
                    verticalAnchor = VerticalAnchor.VerticalAnchorMixed;
                    break;
                default:
                    verticalAnchor = VerticalAnchor.AnchorTop;
                    break;
            }
            return verticalAnchor;
        }

        public static MsoVerticalAnchor ConvertBack(VerticalAnchor verticalAnchor)
        {
            MsoVerticalAnchor msoVerticalAnchor;
            switch (verticalAnchor)
            {
                case VerticalAnchor.AnchorBottom:
                    msoVerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                    break;
                case VerticalAnchor.AnchorBottomBaseLine:
                    msoVerticalAnchor = MsoVerticalAnchor.msoAnchorBottomBaseLine;
                    break;
                case VerticalAnchor.AnchorMiddle:
                    msoVerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    break;
                case VerticalAnchor.AnchorTopBaseline:
                    msoVerticalAnchor = MsoVerticalAnchor.msoAnchorTopBaseline;
                    break;
                case VerticalAnchor.VerticalAnchorMixed:
                    msoVerticalAnchor = MsoVerticalAnchor.msoVerticalAnchorMixed;
                    break;
                default:
                    msoVerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                    break;
            }
            return msoVerticalAnchor;
        }
    }
}
