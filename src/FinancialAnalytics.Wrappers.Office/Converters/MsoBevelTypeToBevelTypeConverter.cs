using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoBevelTypeToBevelTypeConverter
	{
		public static BevelType Convert(MsoBevelType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoBevelType, BevelType>(msoControlType);
		}

		public static MsoBevelType ConvertBack(BevelType controlType)
		{
			return MsoEnumConverter.ConvertToMso<BevelType, MsoBevelType>(controlType);
		}
	}
}
