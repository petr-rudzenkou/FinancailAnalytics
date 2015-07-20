using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetMaterialToPresetMaterialConverter
	{
		public static PresetMaterial Convert(MsoPresetMaterial msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetMaterial, PresetMaterial>(msoControlType);
		}

		public static MsoPresetMaterial ConvertBack(PresetMaterial controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetMaterial, MsoPresetMaterial>(controlType);
		}
	}
}
