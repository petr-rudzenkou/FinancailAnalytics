using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetLightingSoftnessToPresetLightingSoftnessConverter
	{
		public static PresetLightingSoftness Convert(MsoPresetLightingSoftness msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetLightingSoftness, PresetLightingSoftness>(msoControlType);
		}

		public static MsoPresetLightingSoftness ConvertBack(PresetLightingSoftness controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetLightingSoftness, MsoPresetLightingSoftness>(controlType);
		}
	}
}
