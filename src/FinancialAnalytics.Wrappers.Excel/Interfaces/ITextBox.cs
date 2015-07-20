using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface ITextBox : IEntityWrapper<ITextBox>
	{
		IInterior Interior { get; }
		IBorder Border { get; }
		IFont Font { get; }
	}
}
