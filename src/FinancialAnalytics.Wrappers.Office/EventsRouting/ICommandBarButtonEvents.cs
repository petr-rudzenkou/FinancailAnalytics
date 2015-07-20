using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office.EventsRouting
{
	[ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), Guid(CommandBarButton.CommandBarButtonEventsGuid)]
	public interface ICommandBarButtonEvents
	{
		[DispId(0x00000001)]
		void OnClick(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault);
	}
}
