using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IColorStop : IEntityWrapper<IColorStop>
	{
		object Color { get; set; }
		int ThemeColor { get; set; }
		double Position { get; set; }
	}
}