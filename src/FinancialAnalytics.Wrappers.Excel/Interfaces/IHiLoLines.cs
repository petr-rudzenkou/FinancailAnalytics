using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IHiLoLines : IEntityWrapper<IHiLoLines>
	{
		string Name { get; }
		IBorder Border { get; }
	}
}
