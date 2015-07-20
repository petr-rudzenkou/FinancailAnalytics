using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetExtrusionDirectionToPresetExtrusionDirectionConverter
	{
		public static PresetExtrusionDirection Convert(MsoPresetExtrusionDirection msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetExtrusionDirection, PresetExtrusionDirection>(msoControlType);
		}

		public static MsoPresetExtrusionDirection ConvertBack(PresetExtrusionDirection controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetExtrusionDirection, MsoPresetExtrusionDirection>(controlType);
		}
	}
}
