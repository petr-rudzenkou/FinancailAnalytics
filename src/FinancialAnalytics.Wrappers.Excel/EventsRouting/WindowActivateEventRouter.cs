using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WindowActivateEventHandler(IWorkbook workbook, IWindow window);

    internal class WindowActivateEventRouter : ExcelBaseEvent
    {
        public WindowActivateEventRouter()
            : base(0x00000614)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WindowActivateEventHandler(WorkbookWindowPrivateHandler);
            PublicEvent = new WindowActivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook, IWindow window)
        {

        }
    }
}
