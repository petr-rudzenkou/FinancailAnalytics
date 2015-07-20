using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoShadowStyleToShadowStyleConverter
	{
		public static ShadowStyle Convert(MsoShadowStyle msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoShadowStyle, ShadowStyle>(msoControlType);
		}

		public static MsoShadowStyle ConvertBack(ShadowStyle controlType)
		{
			return MsoEnumConverter.ConvertToMso<ShadowStyle, MsoShadowStyle>(controlType);
		}
	}
}
