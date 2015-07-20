using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookBeforePrintEventHandler(IWorkbook workbook, ref bool cancel);

    internal class WorkbookBeforePrintEventRouter : ExcelBaseEvent
    {
        public WorkbookBeforePrintEventRouter()
            : base(0x00000624)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforePrintEventHandler(WorkbookRefBoolPrivateHandler);
            PublicEvent = new WorkbookBeforePrintEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook, ref bool cancel)
        {

        }
    }
}
