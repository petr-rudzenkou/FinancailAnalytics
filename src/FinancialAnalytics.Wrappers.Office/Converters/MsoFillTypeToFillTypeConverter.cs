using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoFillTypeToFillTypeConverter
    {
        public static FillType Convert(MsoFillType msoFillType)
        {
            FillType fillType;
            switch (msoFillType)
            {
                case MsoFillType.msoFillTextured:
                    fillType = FillType.FillTextured;
                    break;
                case MsoFillType.msoFillGradient:
                    fillType = FillType.FillGradient;
                    break;
                case MsoFillType.msoFillMixed:
                    fillType = FillType.FillMixed;
                    break;
                case MsoFillType.msoFillPatterned:
                    fillType = FillType.FillPatterned;
                    break;
                case MsoFillType.msoFillPicture:
                    fillType = FillType.FillPicture;
                    break;
                case MsoFillType.msoFillSolid:
                    fillType = FillType.FillSolid;
                    break;
                case MsoFillType.msoFillBackground:
                    fillType = FillType.FillBackground;
                    break;
                default:
                    fillType = FillType.FillSolid;
                    break;
            }
            return fillType;
        }

        public static MsoFillType ConvertBack(FillType fillType)
        {
            MsoFillType msoFillType;
            switch (fillType)
            {
                case FillType.FillTextured:
                    msoFillType = MsoFillType.msoFillTextured;
                    break;
                case FillType.FillGradient:
                    msoFillType = MsoFillType.msoFillGradient;
                    break;
                case FillType.FillMixed:
                    msoFillType = MsoFillType.msoFillMixed;
                    break;
                case FillType.FillPatterned:
                    msoFillType = MsoFillType.msoFillPatterned;
                    break;
                case FillType.FillPicture:
                    msoFillType = MsoFillType.msoFillPicture;
                    break;
                case FillType.FillSolid:
                    msoFillType = MsoFillType.msoFillSolid;
                    break;
                default:
                    msoFillType = MsoFillType.msoFillBackground;
                    break;
            }
            return msoFillType;
        }
    }
}
