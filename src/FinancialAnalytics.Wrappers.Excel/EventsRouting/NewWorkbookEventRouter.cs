using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void NewWorkbookEventHandler(IWorkbook workbook);

    internal class NewWorkbookEventRouter : ExcelBaseEvent
    {
        public NewWorkbookEventRouter()
            : base(0x0000061D)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_NewWorkbookEventHandler(WorkbookPrivateHandler);
            PublicEvent = new NewWorkbookEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
