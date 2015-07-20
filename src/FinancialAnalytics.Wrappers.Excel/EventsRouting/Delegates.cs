using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{

	public delegate void UnsupportedWorkbookChangeEventHandler(IWorkbook oldWorkbook, IWorkbook newWorkbook);

	public delegate void ChartSelectEventHandler(int elementId, int arg1, int arg2);

	public delegate void ChartMouseDownEventHandler(int button, int shift, int x, int y);
}
