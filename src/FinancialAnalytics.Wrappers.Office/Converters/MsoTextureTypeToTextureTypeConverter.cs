using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoTextureTypeToTextureTypeConverter
	{
		public static TextureType Convert(MsoTextureType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoTextureType, TextureType>(msoControlType);
		}

		public static MsoTextureType ConvertBack(TextureType controlType)
		{
			return MsoEnumConverter.ConvertToMso<TextureType, MsoTextureType>(controlType);
		}
	}
}
