using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPresetCameraToPresetCameraConverter
	{
		public static PresetCamera Convert(MsoPresetCamera msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPresetCamera, PresetCamera>(msoControlType);
		}

		public static MsoPresetCamera ConvertBack(PresetCamera controlType)
		{
			return MsoEnumConverter.ConvertToMso<PresetCamera, MsoPresetCamera>(controlType);
		}
	}
}
