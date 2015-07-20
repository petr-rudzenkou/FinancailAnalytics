using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookNewSheetEventHandler(IWorkbook workbook, ISheet sheet);

    internal class WorkbookNewSheetEventRouter : ExcelBaseEvent
    {
        public WorkbookNewSheetEventRouter()
            : base(0x00000625)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookNewSheetEventHandler(PrivateHandler);
            PublicEvent = new WorkbookNewSheetEventHandler(PublicDefaultHandler);
        }

        private void PrivateHandler(Microsoft.Office.Interop.Excel.Workbook excelWorkbook, object excelSheet)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook;
                ISheet sheet;
                lock (Lock)
                {
                    workbook = EntityResolver.ResolveWorkbook(excelWorkbook);
                    sheet = EntityResolver.ResolveSheet(excelSheet);
                }
                PublicEvent.DynamicInvoke(workbook, sheet);
            }
        }

        private void PublicDefaultHandler(IWorkbook workbook, ISheet sheet)
        {

        }
    }
}
