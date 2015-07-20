using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IChartGroup : IEntityWrapper<IChartGroup>
	{
		bool HasHiLoLines { get; set; }
		bool HasUpDownBars { get; set; }
		IBars DownBars { get; }
		IBars UpBars { get; }
		IHiLoLines HiLoLines { get; }
	}
}
