using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoAutoShapeTypeToAutoShapeTypeConverter
	{
		public static AutoShapeType Convert(MsoAutoShapeType msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoAutoShapeType, AutoShapeType>(msoControlType);
		}

		public static MsoAutoShapeType ConvertBack(AutoShapeType controlType)
		{
			return MsoEnumConverter.ConvertToMso<AutoShapeType, MsoAutoShapeType>(controlType);
		}
	}
}