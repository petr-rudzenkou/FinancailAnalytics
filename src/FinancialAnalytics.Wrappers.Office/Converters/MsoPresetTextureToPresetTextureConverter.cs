using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoPresetTextureToPresetTextureConverter
    {
        public static PresetTexture Convert(MsoPresetTexture msoPresetTexture)
        {
            PresetTexture presetTexture;
            switch (msoPresetTexture)
            {
                case MsoPresetTexture.msoPresetTextureMixed:
                    presetTexture = PresetTexture.PresetTextureMixed;
                    break;
                case MsoPresetTexture.msoTextureBlueTissuePaper:
                    presetTexture = PresetTexture.TextureBlueTissuePaper;
                    break;
                case MsoPresetTexture.msoTextureBouquet:
                    presetTexture = PresetTexture.TextureBouquet;
                    break;
                case MsoPresetTexture.msoTextureBrownMarble:
                    presetTexture = PresetTexture.TextureBrownMarble;
                    break;
                case MsoPresetTexture.msoTextureCanvas:
                    presetTexture = PresetTexture.TextureCanvas;
                    break;
                case MsoPresetTexture.msoTextureCork:
                    presetTexture = PresetTexture.TextureCork;
                    break;
                case MsoPresetTexture.msoTextureDenim:
                    presetTexture = PresetTexture.TextureDenim;
                    break;
                case MsoPresetTexture.msoTextureFishFossil:
                    presetTexture = PresetTexture.TextureFishFossil;
                    break;
                case MsoPresetTexture.msoTextureGranite:
                    presetTexture = PresetTexture.TextureGranite;
                    break;
                case MsoPresetTexture.msoTextureGreenMarble:
                    presetTexture = PresetTexture.TextureGreenMarble;
                    break;
                case MsoPresetTexture.msoTextureMediumWood:
                    presetTexture = PresetTexture.TextureMediumWood;
                    break;
                case MsoPresetTexture.msoTextureNewsprint:
                    presetTexture = PresetTexture.TextureNewsprint;
                    break;
                case MsoPresetTexture.msoTextureOak:
                    presetTexture = PresetTexture.TextureOak;
                    break;
                case MsoPresetTexture.msoTexturePaperBag:
                    presetTexture = PresetTexture.TexturePaperBag;
                    break;
                case MsoPresetTexture.msoTexturePapyrus:
                    presetTexture = PresetTexture.TexturePapyrus;
                    break;
                case MsoPresetTexture.msoTextureParchment:
                    presetTexture = PresetTexture.TextureParchment;
                    break;
                case MsoPresetTexture.msoTexturePinkTissuePaper:
                    presetTexture = PresetTexture.TexturePinkTissuePaper;
                    break;
                case MsoPresetTexture.msoTexturePurpleMesh:
                    presetTexture = PresetTexture.TexturePurpleMesh;
                    break;
                case MsoPresetTexture.msoTextureRecycledPaper:
                    presetTexture = PresetTexture.TextureRecycledPaper;
                    break;
                case MsoPresetTexture.msoTextureSand:
                    presetTexture = PresetTexture.TextureSand;
                    break;
                case MsoPresetTexture.msoTextureStationery:
                    presetTexture = PresetTexture.TextureStationery;
                    break;
                case MsoPresetTexture.msoTextureWalnut:
                    presetTexture = PresetTexture.TextureWalnut;
                    break;
                case MsoPresetTexture.msoTextureWaterDroplets:
                    presetTexture = PresetTexture.TextureWaterDroplets;
                    break;
                case MsoPresetTexture.msoTextureWhiteMarble:
                    presetTexture = PresetTexture.TextureWhiteMarble;
                    break;
                case MsoPresetTexture.msoTextureWovenMat:
                    presetTexture = PresetTexture.TextureWovenMat;
                    break;
                default:
                    presetTexture = PresetTexture.PresetTextureMixed;
                    break;
            }
            return presetTexture;
        }
        public static MsoPresetTexture ConvertBack(PresetTexture presetTexture)
        {
            MsoPresetTexture msoPresetTexture;
            switch (presetTexture)
            {
                case PresetTexture.PresetTextureMixed:
                    msoPresetTexture = MsoPresetTexture.msoPresetTextureMixed;
                    break;
                case PresetTexture.TextureBlueTissuePaper:
                    msoPresetTexture = MsoPresetTexture.msoTextureBlueTissuePaper;
                    break;
                case PresetTexture.TextureBouquet:
                    msoPresetTexture = MsoPresetTexture.msoTextureBouquet;
                    break;
                case PresetTexture.TextureBrownMarble:
                    msoPresetTexture = MsoPresetTexture.msoTextureBrownMarble;
                    break;
                case PresetTexture.TextureCanvas:
                    msoPresetTexture = MsoPresetTexture.msoTextureCanvas;
                    break;
                case PresetTexture.TextureCork:
                    msoPresetTexture = MsoPresetTexture.msoTextureCork;
                    break;
                case PresetTexture.TextureDenim:
                    msoPresetTexture = MsoPresetTexture.msoTextureDenim;
                    break;
                case PresetTexture.TextureFishFossil:
                    msoPresetTexture = MsoPresetTexture.msoTextureFishFossil;
                    break;
                case PresetTexture.TextureGranite:
                    msoPresetTexture = MsoPresetTexture.msoTextureGranite;
                    break;
                case PresetTexture.TextureGreenMarble:
                    msoPresetTexture = MsoPresetTexture.msoTextureGreenMarble;
                    break;
                case PresetTexture.TextureMediumWood:
                    msoPresetTexture = MsoPresetTexture.msoTextureMediumWood;
                    break;
                case PresetTexture.TextureNewsprint:
                    msoPresetTexture = MsoPresetTexture.msoTextureNewsprint;
                    break;
                case PresetTexture.TextureOak:
                    msoPresetTexture = MsoPresetTexture.msoTextureOak;
                    break;
                case PresetTexture.TexturePaperBag:
                    msoPresetTexture = MsoPresetTexture.msoTexturePaperBag;
                    break;
                case PresetTexture.TexturePapyrus:
                    msoPresetTexture = MsoPresetTexture.msoTexturePapyrus;
                    break;
                case PresetTexture.TextureParchment:
                    msoPresetTexture = MsoPresetTexture.msoTextureParchment;
                    break;
                case PresetTexture.TexturePinkTissuePaper:
                    msoPresetTexture = MsoPresetTexture.msoTexturePinkTissuePaper;
                    break;
                case PresetTexture.TexturePurpleMesh:
                    msoPresetTexture = MsoPresetTexture.msoTexturePurpleMesh;
                    break;
                case PresetTexture.TextureRecycledPaper:
                    msoPresetTexture = MsoPresetTexture.msoTextureRecycledPaper;
                    break;
                case PresetTexture.TextureSand:
                    msoPresetTexture = MsoPresetTexture.msoTextureSand;
                    break;
                case PresetTexture.TextureStationery:
                    msoPresetTexture = MsoPresetTexture.msoTextureStationery;
                    break;
                case PresetTexture.TextureWalnut:
                    msoPresetTexture = MsoPresetTexture.msoTextureWalnut;
                    break;
                case PresetTexture.TextureWaterDroplets:
                    msoPresetTexture = MsoPresetTexture.msoTextureWaterDroplets;
                    break;
                case PresetTexture.TextureWhiteMarble:
                    msoPresetTexture = MsoPresetTexture.msoTextureWhiteMarble;
                    break;
                case PresetTexture.TextureWovenMat:
                    msoPresetTexture = MsoPresetTexture.msoTextureWovenMat;
                    break;
                default:
                    msoPresetTexture = MsoPresetTexture.msoPresetTextureMixed;
                    break;
            }
            return msoPresetTexture;
        }
    }
}
