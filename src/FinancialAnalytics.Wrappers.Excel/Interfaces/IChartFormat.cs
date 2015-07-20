using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{

	public interface IChartFormat : IEntityWrapper<IChartFormat>
	{
		ITextFrame2 TextFrame2 { get; }
		IFillFormat Fill { get; }
		ILineFormat Line { get; }
	}
}
