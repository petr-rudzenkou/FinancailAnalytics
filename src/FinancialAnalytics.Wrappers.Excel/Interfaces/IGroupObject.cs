using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IGroupObject : IEntityWrapper<IGroupObject>
	{
		IInterior Interior { get; }
		IBorder Border { get; }
		IFont Font { get; }
	}
}
