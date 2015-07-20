using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoColorTypeToColorTypeConverter
	{
		public static ColorType Convert(MsoColorType msoColorType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoColorType, ColorType>(msoColorType);
		}

		public static MsoColorType ConvertBack(ColorType colorType)
		{
			return MsoEnumConverter.ConvertToMso<ColorType, MsoColorType>(colorType);
		}
	}
}