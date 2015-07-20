using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{

	public class XlDisplayUnitToDisplayUnitConverter
	{
		public static DisplayUnit Convert(XlDisplayUnit xlDisplayUnit)
		{
			DisplayUnit displayUnit;
			switch (xlDisplayUnit)
			{
				case XlDisplayUnit.xlHundredMillions:
					displayUnit = DisplayUnit.HundredMillions;
					break;
				case XlDisplayUnit.xlHundreds:
					displayUnit = DisplayUnit.Hundreds;
					break;
				case XlDisplayUnit.xlHundredThousands:
					displayUnit = DisplayUnit.HundredThousands;
					break;
				case XlDisplayUnit.xlMillionMillions:
					displayUnit = DisplayUnit.MillionMillions;
					break;
				case XlDisplayUnit.xlMillions:
					displayUnit = DisplayUnit.Millions;
					break;
				case XlDisplayUnit.xlTenMillions:
					displayUnit = DisplayUnit.TenMillions;
					break;
				case XlDisplayUnit.xlTenThousands:
					displayUnit = DisplayUnit.TenThousands;
					break;
				case XlDisplayUnit.xlThousandMillions:
					displayUnit = DisplayUnit.ThousandMillions;
					break;
				case XlDisplayUnit.xlThousands:
					displayUnit = DisplayUnit.Thousands;
					break;
				default:
					displayUnit = DisplayUnit.IsNotSet;
					break;
			}
			return displayUnit;
		}

		public static XlDisplayUnit ConvertBack(DisplayUnit displayUnit)
		{
			XlDisplayUnit xlDisplayUnit;
			switch (displayUnit)
			{
				case DisplayUnit.HundredMillions:
					xlDisplayUnit = XlDisplayUnit.xlHundredMillions;
					break;
				case DisplayUnit.Hundreds:
					xlDisplayUnit = XlDisplayUnit.xlHundreds;
					break;
				case DisplayUnit.HundredThousands:
					xlDisplayUnit = XlDisplayUnit.xlHundredThousands;
					break;
				case DisplayUnit.MillionMillions:
					xlDisplayUnit = XlDisplayUnit.xlMillionMillions;
					break;
				case DisplayUnit.Millions:
					xlDisplayUnit = XlDisplayUnit.xlMillions;
					break;
				case DisplayUnit.TenMillions:
					xlDisplayUnit = XlDisplayUnit.xlTenMillions;
					break;
				case DisplayUnit.TenThousands:
					xlDisplayUnit = XlDisplayUnit.xlTenThousands;
					break;
				case DisplayUnit.ThousandMillions:
					xlDisplayUnit = XlDisplayUnit.xlThousandMillions;
					break;
				case DisplayUnit.Thousands:
					xlDisplayUnit = XlDisplayUnit.xlThousands;
					break;
				default:
					xlDisplayUnit = XlDisplayUnit.xlThousands;
					break;
			}
			return xlDisplayUnit;			
		}
	}
}
