using System.Drawing;
using FinancialAnalytics.Wrappers.Office.EventsRouting;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICommandBarButton : ICommandBarControl, ICommandBarButtonEvents
	{
		event _CommandBarButtonEvents_ClickEventHandler Click;

		Bitmap Picture { get; set; }

		Bitmap Mask { get; set; }

		CommandBarButtonState State { get; set; }

		void Execute();

		void TurnOnButton();

		void TurnOffButton();
	}
}