using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IColorFormat
	{
		int RGB { get; set; }
		ColorType Type { get; }
	}
}
