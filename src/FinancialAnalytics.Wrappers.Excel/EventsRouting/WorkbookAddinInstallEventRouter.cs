using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookAddinInstallEventHandler(IWorkbook workbook);

    internal class WorkbookAddinInstallEventRouter : ExcelBaseEvent
    {
        public WorkbookAddinInstallEventRouter()
            : base(0x00000626)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAddinInstallEventHandler(WorkbookPrivateHandler);
            PublicEvent = new WorkbookAddinInstallEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
