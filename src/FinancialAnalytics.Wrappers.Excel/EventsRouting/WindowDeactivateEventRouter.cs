using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="workbook">The workbook displayed in the deactivated window.</param>
    /// <param name="window">The deactivated window.</param>
    public delegate void WindowDeactivateEventHandler(IWorkbook workbook, IWindow window);

    internal class WindowDeactivateEventRouter : ExcelBaseEvent
    {
        public WindowDeactivateEventRouter()
            : base(0x00000615)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WindowDeactivateEventHandler(WorkbookWindowPrivateHandler);
            PublicEvent = new WindowDeactivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook, IWindow window)
        {

        }
    }
}
