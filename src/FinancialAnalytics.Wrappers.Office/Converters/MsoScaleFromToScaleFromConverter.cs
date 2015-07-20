using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{

    public class MsoScaleFromToScaleFromConverter
    {
        public static ScaleFrom Convert(MsoScaleFrom msoScaleFrom)
        {
            ScaleFrom scaleFrom = ScaleFrom.ScaleFromMiddle;
            switch (msoScaleFrom)
            {
                case MsoScaleFrom.msoScaleFromBottomRight:
                    scaleFrom = ScaleFrom.ScaleFromBottomRight;
                    break;
                case MsoScaleFrom.msoScaleFromTopLeft:
                    scaleFrom = ScaleFrom.ScaleFromTopLeft;
                    break;
                default:
                    scaleFrom = ScaleFrom.ScaleFromMiddle;
                    break;
            }
            return scaleFrom;
        }

        public static MsoScaleFrom ConvertBack(ScaleFrom scaleFrom)
        {
            MsoScaleFrom msoScaleFrom = MsoScaleFrom.msoScaleFromMiddle;
            switch (scaleFrom)
            {
                case ScaleFrom.ScaleFromBottomRight:
                    msoScaleFrom = MsoScaleFrom.msoScaleFromBottomRight;
                    break;
                case ScaleFrom.ScaleFromTopLeft:
                    msoScaleFrom = MsoScaleFrom.msoScaleFromTopLeft;
                    break;
                default:
                    msoScaleFrom = MsoScaleFrom.msoScaleFromMiddle;
                    break;
            }
            return msoScaleFrom;
        }
    }
}
