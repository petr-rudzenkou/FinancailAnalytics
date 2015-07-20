using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoLightRigTypeToLightRigTypeConverter
	{
		public static LightRigType Convert(MsoLightRigType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoLightRigType, LightRigType>(msoControlType);
		}

		public static MsoLightRigType ConvertBack(LightRigType controlType)
		{
			return MsoEnumConverter.ConvertToMso<LightRigType, MsoLightRigType>(controlType);
		}
	}
}
