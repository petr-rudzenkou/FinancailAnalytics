using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoPictureColorTypeToPictureColorTypeConverter
    {
        public static PictureColorType Convert(MsoPictureColorType msoPictureColorType)
        {
            PictureColorType pictureColorType;
            switch (msoPictureColorType)
            {
                case MsoPictureColorType.msoPictureBlackAndWhite:
                    pictureColorType = PictureColorType.PictureBlackAndWhite;
                    break;
                case MsoPictureColorType.msoPictureGrayscale:
                    pictureColorType = PictureColorType.PictureGrayscale;
                    break;
                case MsoPictureColorType.msoPictureMixed:
                    pictureColorType = PictureColorType.PictureMixed;
                    break;
                case MsoPictureColorType.msoPictureWatermark:
                    pictureColorType = PictureColorType.PictureWatermark;
                    break;
                default:
                    pictureColorType = PictureColorType.PictureAutomatic;
                    break;
            }
            return pictureColorType;
        }

        public static MsoPictureColorType ConvertBack(PictureColorType pictureColorType)
        {
            MsoPictureColorType msoPictureColorType;
            switch (pictureColorType)
            {
                case PictureColorType.PictureBlackAndWhite:
                    msoPictureColorType = MsoPictureColorType.msoPictureBlackAndWhite;
                    break;
                case PictureColorType.PictureGrayscale:
                    msoPictureColorType = MsoPictureColorType.msoPictureGrayscale;
                    break;
                case PictureColorType.PictureMixed:
                    msoPictureColorType = MsoPictureColorType.msoPictureMixed;
                    break;
                case PictureColorType.PictureWatermark:
                    msoPictureColorType = MsoPictureColorType.msoPictureWatermark;
                    break;
                default:
                    msoPictureColorType = MsoPictureColorType.msoPictureAutomatic;
                    break;
            }
            return msoPictureColorType;
        }
    }
}
