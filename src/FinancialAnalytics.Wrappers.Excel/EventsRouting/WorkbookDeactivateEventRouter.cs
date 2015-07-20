using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookDeactivateEventHandler(IWorkbook workbook);

    internal class WorkbookDeactivateEventRouter : ExcelBaseEvent
    {
        public WorkbookDeactivateEventRouter()
            : base(0x00000621)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookDeactivateEventHandler(WorkbookPrivateHandler);
            PublicEvent = new WorkbookDeactivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
