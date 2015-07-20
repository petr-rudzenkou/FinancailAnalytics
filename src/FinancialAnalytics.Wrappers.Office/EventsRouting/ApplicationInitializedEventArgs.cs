using System;

namespace FinancialAnalytics.Wrappers.Office.EventsRouting
{
	public class ApplicationInitializedEventArgs : EventArgs
	{
		public ApplicationInitializedEventArgs(int applicationWindowHandle)
		{
			ApplicationWindowHandle = applicationWindowHandle;
		}

		public int ApplicationWindowHandle { get; private set; }
	}
}
