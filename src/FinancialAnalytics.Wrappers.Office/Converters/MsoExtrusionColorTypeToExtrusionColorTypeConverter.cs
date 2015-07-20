using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoExtrusionColorTypeToExtrusionColorTypeConverter
	{
		public static ExtrusionColorType Convert(MsoExtrusionColorType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoExtrusionColorType, ExtrusionColorType>(msoControlType);
		}

		public static MsoExtrusionColorType ConvertBack(ExtrusionColorType controlType)
		{
			return MsoEnumConverter.ConvertToMso<ExtrusionColorType, MsoExtrusionColorType>(controlType);
		}
	}
}
