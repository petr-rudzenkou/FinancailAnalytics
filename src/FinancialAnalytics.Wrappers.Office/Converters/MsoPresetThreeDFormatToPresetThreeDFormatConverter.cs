using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetThreeDFormatToPresetThreeDFormatConverter
	{
		public static PresetThreeDFormat Convert(MsoPresetThreeDFormat msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetThreeDFormat, PresetThreeDFormat>(msoControlType);
		}

		public static MsoPresetThreeDFormat ConvertBack(PresetThreeDFormat controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetThreeDFormat, MsoPresetThreeDFormat>(controlType);
		}
	}
}
