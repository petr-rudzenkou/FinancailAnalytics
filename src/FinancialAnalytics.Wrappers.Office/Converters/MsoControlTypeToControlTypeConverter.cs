using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoControlTypeToControlTypeConverter
	{
		public static ControlType Convert(MsoControlType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoControlType, ControlType>(msoControlType);
		}

		public static MsoControlType ConvertBack(ControlType controlType)
		{
			return MsoEnumConverter.ConvertToMso<ControlType, MsoControlType>(controlType);
		}
	}
}