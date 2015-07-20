using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoArrowheadLengthToArrowheadLengthConverter
	{
		public static ArrowheadLength Convert(MsoArrowheadLength msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoArrowheadLength, ArrowheadLength>(msoControlType);
		}

		public static MsoArrowheadLength ConvertBack(ArrowheadLength controlType)
		{
			return MsoEnumConverter.ConvertToMso<ArrowheadLength, MsoArrowheadLength>(controlType);
		}
	}
}
