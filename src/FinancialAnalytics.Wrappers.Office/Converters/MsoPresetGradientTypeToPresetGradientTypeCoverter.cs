using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoPresetGradientTypeToPresetGradientTypeCoverter
    {
        public static PresetGradientType Convert (MsoPresetGradientType msoPresetGradientType)
        {
            PresetGradientType presetGradientType;
            switch (msoPresetGradientType)
            {
                case MsoPresetGradientType.msoGradientBrass :
                    presetGradientType = PresetGradientType.GradientBrass;
                    break;
                case MsoPresetGradientType.msoGradientCalmWater :
                    presetGradientType = PresetGradientType.GradientCalmWater;
                    break;
                case MsoPresetGradientType.msoGradientChrome :
                    presetGradientType = PresetGradientType.GradientChrome;
                    break;
                case MsoPresetGradientType.msoGradientChromeII :
                    presetGradientType = PresetGradientType.GradientChromeII;
                    break;
                case MsoPresetGradientType.msoGradientDaybreak :
                    presetGradientType = PresetGradientType.GradientDaybreak;
                    break;
                case MsoPresetGradientType.msoGradientDesert :
                    presetGradientType = PresetGradientType.GradientDesert;
                    break;
                case MsoPresetGradientType.msoGradientEarlySunset :
                    presetGradientType = PresetGradientType.GradientEarlySunset;
                    break;
                case MsoPresetGradientType.msoGradientFire :
                    presetGradientType = PresetGradientType.GradientFire;
                    break;
                case MsoPresetGradientType.msoGradientFog :
                    presetGradientType = PresetGradientType.GradientFog;
                    break;
                case MsoPresetGradientType.msoGradientGold :
                    presetGradientType = PresetGradientType.GradientGold;
                    break;
                case MsoPresetGradientType.msoGradientGoldII :
                    presetGradientType = PresetGradientType.GradientGoldII;
                    break;
                case MsoPresetGradientType.msoGradientHorizon :
                    presetGradientType = PresetGradientType.GradientHorizon;
                    break;
                case MsoPresetGradientType.msoGradientLateSunset :
                    presetGradientType = PresetGradientType.GradientLateSunset;
                    break;
                case MsoPresetGradientType.msoGradientMahogany :
                    presetGradientType = PresetGradientType.GradientMahogany;
                    break;
                case MsoPresetGradientType.msoGradientMoss :
                    presetGradientType = PresetGradientType.GradientMoss;
                    break;
                case MsoPresetGradientType.msoGradientNightfall :
                    presetGradientType = PresetGradientType.GradientNightfall;
                    break;
                case MsoPresetGradientType.msoGradientOcean :
                    presetGradientType = PresetGradientType.GradientOcean;
                    break;
                case MsoPresetGradientType.msoGradientParchment :
                    presetGradientType = PresetGradientType.GradientParchment;
                    break;
                case MsoPresetGradientType.msoGradientPeacock :
                    presetGradientType = PresetGradientType.GradientPeacock;
                    break;
                case MsoPresetGradientType.msoGradientRainbow :
                    presetGradientType = PresetGradientType.GradientRainbow;
                    break;
                case MsoPresetGradientType.msoGradientRainbowII :
                    presetGradientType = PresetGradientType.GradientRainbowII;
                    break;
                case MsoPresetGradientType.msoGradientSapphire :
                    presetGradientType = PresetGradientType.GradientSapphire;
                    break;
                case MsoPresetGradientType.msoGradientSilver :
                    presetGradientType = PresetGradientType.GradientSilver;
                    break;
                case MsoPresetGradientType.msoGradientWheat :
                    presetGradientType = PresetGradientType.GradientWheat;
                    break;
                default :
                    presetGradientType = PresetGradientType.PresetGradientMixed;
                    break;
            }
            return presetGradientType;
        }

        public static MsoPresetGradientType ConvertBack(PresetGradientType presetGradientType)
        {
            MsoPresetGradientType msoPresetGradientType;
            switch (presetGradientType)
            {
                case PresetGradientType.GradientBrass:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientBrass;
                    break;
                case PresetGradientType.GradientCalmWater:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientCalmWater;
                    break;
                case PresetGradientType.GradientChrome:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientChrome;
                    break;
                case PresetGradientType.GradientChromeII:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientChromeII;
                    break;
                case PresetGradientType.GradientDaybreak:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientDaybreak;
                    break;
                case PresetGradientType.GradientDesert:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientDesert;
                    break;
                case PresetGradientType.GradientEarlySunset:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientEarlySunset;
                    break;
                case PresetGradientType.GradientFire:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientFire;
                    break;
                case PresetGradientType.GradientFog:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientFog;
                    break;
                case PresetGradientType.GradientGold:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientGold;
                    break;
                case PresetGradientType.GradientGoldII:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientGoldII;
                    break;
                case PresetGradientType.GradientHorizon:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientHorizon;
                    break;
                case PresetGradientType.GradientLateSunset:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientLateSunset;
                    break;
                case PresetGradientType.GradientMahogany:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientMahogany;
                    break;
                case PresetGradientType.GradientMoss:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientMoss;
                    break;
                case PresetGradientType.GradientNightfall:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientNightfall;
                    break;
                case PresetGradientType.GradientOcean:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientOcean;
                    break;
                case PresetGradientType.GradientParchment:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientParchment;
                    break;
                case PresetGradientType.GradientPeacock:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientPeacock;
                    break;
                case PresetGradientType.GradientRainbow:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientRainbow;
                    break;
                case PresetGradientType.GradientRainbowII:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientRainbowII;
                    break;
                case PresetGradientType.GradientSapphire:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientSapphire;
                    break;
                case PresetGradientType.GradientSilver:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientSilver;
                    break;
                case PresetGradientType.GradientWheat:
                    msoPresetGradientType = MsoPresetGradientType.msoGradientWheat;
                    break;
                default:
                    msoPresetGradientType = MsoPresetGradientType.msoPresetGradientMixed;
                    break;
            }
            return msoPresetGradientType;
        }
    }
}
