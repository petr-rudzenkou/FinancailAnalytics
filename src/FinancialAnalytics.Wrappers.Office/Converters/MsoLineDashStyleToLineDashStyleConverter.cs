using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoLineDashStyleToLineDashStyleConverter
	{
		public static LineDashStyle Convert(MsoLineDashStyle msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoLineDashStyle, LineDashStyle>(msoControlType);
		}

		public static MsoLineDashStyle ConvertBack(LineDashStyle controlType)
		{
			return MsoEnumConverter.ConvertToMso<LineDashStyle, MsoLineDashStyle>(controlType);
		}
	}
}
