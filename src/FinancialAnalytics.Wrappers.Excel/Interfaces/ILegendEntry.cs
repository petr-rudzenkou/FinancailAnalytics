using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface ILegendEntry : IEntityWrapper<ILegendEntry>
	{
		Wrappers.Excel.Interfaces.IFont Font { get; }

		Wrappers.Excel.Interfaces.ILegendKey LegendKey { get; }
	}
}
