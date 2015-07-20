using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WindowResizeEventHandler(IWorkbook workbook, IWindow window);

    internal class WindowResizeEventRouter : ExcelBaseEvent
    {
        public WindowResizeEventRouter()
            : base(0x00000612)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WindowResizeEventHandler(WorkbookWindowPrivateHandler);
            PublicEvent = new WindowResizeEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook, IWindow window)
        {
            
        }
    }
}
