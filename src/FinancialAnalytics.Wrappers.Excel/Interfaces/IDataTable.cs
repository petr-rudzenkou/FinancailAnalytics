using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IDataTable : IEntityWrapper<IDataTable>
	{
		IBorder Border { get; }
		IFont Font { get; }
        IChartFormat Format { get; }
	}
}
