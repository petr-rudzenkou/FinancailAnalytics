using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoPatternTypeToPatternTypeConverter
	{
		public static PatternType Convert(MsoPatternType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoPatternType, PatternType>(msoControlType);
		}

		public static MsoPatternType ConvertBack(PatternType controlType)
		{
			return MsoEnumConverter.ConvertToMso<PatternType, MsoPatternType>(controlType);
		}
	}
}
