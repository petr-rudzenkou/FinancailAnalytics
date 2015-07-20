using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office.EventsRouting
{
	public delegate void _CommandBarButtonEvents_ClickEventHandler(ICommandBarButton Ctrl, ref bool CancelDefault);

	public delegate void CommandBarsOnUpdate();
}
