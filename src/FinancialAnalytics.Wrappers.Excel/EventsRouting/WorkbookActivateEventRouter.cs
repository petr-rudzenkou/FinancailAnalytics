using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookActivateEventHandler(IWorkbook workbook);

    internal class WorkbookActivateEventRouter : ExcelBaseEvent
    {
        public WorkbookActivateEventRouter()
            : base(0x00000620)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(WorkbookPrivateHandler);
            PublicEvent = new WorkbookActivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
