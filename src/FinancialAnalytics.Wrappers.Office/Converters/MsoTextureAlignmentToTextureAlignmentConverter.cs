using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoTextureAlignmentToTextureAlignmentConverter
	{
		public static TextureAlignment Convert(MsoTextureAlignment msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoTextureAlignment, TextureAlignment>(msoControlType);
		}

		public static MsoTextureAlignment ConvertBack(TextureAlignment controlType)
		{
			return MsoEnumConverter.ConvertToMso<TextureAlignment, MsoTextureAlignment>(controlType);
		}
	}
}
