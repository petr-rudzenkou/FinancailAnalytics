using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IWorksheetFunction : IEntityWrapper<IWorksheetFunction>
	{
		double Max(object range);
		double Match(object value, object range, object index);
		dynamic Index(object range, double value);
		double Min(object range);
	}
}
