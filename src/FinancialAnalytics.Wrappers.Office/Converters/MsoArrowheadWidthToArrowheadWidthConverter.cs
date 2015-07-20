using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoArrowheadWidthToArrowheadWidthConverter
	{
		public static ArrowheadWidth Convert(MsoArrowheadWidth msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoArrowheadWidth, ArrowheadWidth>(msoControlType);
		}

		public static MsoArrowheadWidth ConvertBack(ArrowheadWidth controlType)
		{
			return MsoEnumConverter.ConvertToMso<ArrowheadWidth, MsoArrowheadWidth>(controlType);
		}
	}
}
