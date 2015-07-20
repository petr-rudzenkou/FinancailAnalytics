using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookAddinUninstallEventHandler(IWorkbook workbook);

    internal class WorkbookAddinUninstallEventRouter : ExcelBaseEvent
    {
        public WorkbookAddinUninstallEventRouter()
            : base(0x00000627)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAddinUninstallEventHandler(WorkbookPrivateHandler);
            PublicEvent = new WorkbookAddinUninstallEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook)
        {

        }
    }
}
