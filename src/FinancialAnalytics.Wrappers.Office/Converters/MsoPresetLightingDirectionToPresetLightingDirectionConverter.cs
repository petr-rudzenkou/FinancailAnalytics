using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetLightingDirectionToPresetLightingDirectionConverter
	{
		public static PresetLightingDirection Convert(MsoPresetLightingDirection msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetLightingDirection, PresetLightingDirection>(msoControlType);
		}

		public static MsoPresetLightingDirection ConvertBack(PresetLightingDirection controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetLightingDirection, MsoPresetLightingDirection>(controlType);
		}
	}
}
