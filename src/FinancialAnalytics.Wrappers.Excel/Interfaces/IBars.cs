using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IBars : IEntityWrapper<IBars>
	{
		string Name { get; }
		IBorder Border { get; }
		IInterior Interior { get; }
		IChartFillFormat Fill { get; }
	}
}