using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoGradientColorTypeToGradientColorType
    {
        public static GradientColorType Convert(MsoGradientColorType msoGradientColorType)
        {
            GradientColorType gradientColorType;
            switch (msoGradientColorType)
            {
                case MsoGradientColorType.msoGradientColorMixed:
                    gradientColorType = GradientColorType.GradientColorMixed;
                    break;
                case MsoGradientColorType.msoGradientOneColor:
                    gradientColorType = GradientColorType.GradientOneColor;
                    break;
                case MsoGradientColorType.msoGradientTwoColors:
                    gradientColorType = GradientColorType.GradientTwoColors;
                    break;
                case MsoGradientColorType.msoGradientPresetColors:
                    gradientColorType = GradientColorType.GradientPresetColors;
                    break;
                default:
                    gradientColorType = GradientColorType.GradientMultiColor;
                    break;
            }
            return gradientColorType;
        }

        public static MsoGradientColorType ConvertBack(GradientColorType gradientColorType)
        {
            MsoGradientColorType msoGradientColorType;
            switch (gradientColorType)
            {
                case GradientColorType.GradientColorMixed:
                    msoGradientColorType = MsoGradientColorType.msoGradientColorMixed;
                    break;
                case GradientColorType.GradientOneColor:
                    msoGradientColorType = MsoGradientColorType.msoGradientOneColor;
                    break;
                case GradientColorType.GradientTwoColors:
                    msoGradientColorType = MsoGradientColorType.msoGradientTwoColors;
                    break;
                case GradientColorType.GradientPresetColors:
                    msoGradientColorType = MsoGradientColorType.msoGradientPresetColors;
                    break;
                default:
                    msoGradientColorType = MsoGradientColorType.msoGradientMultiColor;
                    break;
            }
            return msoGradientColorType;
        } 
    }
}
