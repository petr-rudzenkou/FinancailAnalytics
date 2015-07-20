using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookOpenEventHandler(IWorkbook workbook);

    internal class WorkbookOpenEventRouter : ExcelBaseEvent
    {
        public WorkbookOpenEventRouter()
            : base(0x0000061F)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(WorkbookPrivateHandler);
            PublicEvent = new WorkbookOpenEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
