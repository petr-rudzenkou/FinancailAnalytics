using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoTextUnderlineTypeToTextUnderlineTypeConverter
	{
		public static TextUnderlineType Convert(MsoTextUnderlineType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoTextUnderlineType, TextUnderlineType>(msoControlType);
		}

		public static MsoTextUnderlineType ConvertBack(TextUnderlineType controlType)
		{
			return MsoEnumConverter.ConvertToMso<TextUnderlineType, MsoTextUnderlineType>(controlType);
		}
	}
}
