using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoHorizontalAnchorToHorizontalAnchorConverter
    {
        public static HorizontalAnchor Convert(MsoHorizontalAnchor msoHorizontalAnchor)
        {
            HorizontalAnchor horizontalAnchor;
            switch (msoHorizontalAnchor)
            {
                case MsoHorizontalAnchor.msoAnchorCenter:
                    horizontalAnchor = HorizontalAnchor.AnchorCenter;
                    break;
                case MsoHorizontalAnchor.msoHorizontalAnchorMixed:
                    horizontalAnchor = HorizontalAnchor.HorizontalAnchorMixed;
                    break;
                default:
                    horizontalAnchor = HorizontalAnchor.AnchorNone;
                    break;
            }
            return horizontalAnchor;
        }

        public static MsoHorizontalAnchor ConvertBack(HorizontalAnchor horizontalAnchor)
        {
            MsoHorizontalAnchor msoHorizontalAnchor;
            switch (horizontalAnchor)
            {
                case HorizontalAnchor.AnchorCenter:
                    msoHorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
                    break;
                case HorizontalAnchor.HorizontalAnchorMixed:
                    msoHorizontalAnchor = MsoHorizontalAnchor.msoHorizontalAnchorMixed;
                    break;
                default:
                    msoHorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone;
                    break;
            }
            return msoHorizontalAnchor;
        }
    }
}
