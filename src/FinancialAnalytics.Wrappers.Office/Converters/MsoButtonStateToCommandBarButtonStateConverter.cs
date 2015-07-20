using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoButtonStateToCommandBarButtonStateConverter
	{
		public static CommandBarButtonState Convert(MsoButtonState msoControlType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoButtonState, CommandBarButtonState>(msoControlType);
		}

		public static MsoButtonState ConvertBack(CommandBarButtonState controlType)
		{
			return MsoEnumConverter.ConvertToMso<CommandBarButtonState, MsoButtonState>(controlType);
		}
	}
}
