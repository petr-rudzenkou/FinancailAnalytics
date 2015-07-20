using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoArrowheadStyleToArrowheadStyleConverter
	{
		public static ArrowheadStyle Convert(MsoArrowheadStyle msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoArrowheadStyle, ArrowheadStyle>(msoControlType);
		}

		public static MsoArrowheadStyle ConvertBack(ArrowheadStyle controlType)
		{
			return MsoEnumConverter.ConvertToMso<ArrowheadStyle, MsoArrowheadStyle>(controlType);
		}
	}
}
