using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{

	public interface ITextFrame2 : IEntityWrapper<ITextFrame2>
	{
		ITextRange2 TextRange { get; }
	}
}
