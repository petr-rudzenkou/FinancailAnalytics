using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void WorkbookRowsetCompleteEventHandler(IWorkbook workbook, string desciption, string sheet, bool success);

    internal class WorkbookRowsetCompleteEventRouter : ExcelBaseEvent
    {
        public WorkbookRowsetCompleteEventRouter()
            : base(0x00000A33)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookRowsetCompleteEventHandler(PrivateHandler);
            PublicEvent = new WorkbookRowsetCompleteEventHandler(PublicDefaultHandler);
        }

        private void PrivateHandler(Microsoft.Office.Interop.Excel.Workbook excelWorkbook, string description, string sheet, bool success)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook;
                lock (Lock)
                {
                    workbook = EntityResolver.ResolveWorkbook(excelWorkbook);
                }
                PublicEvent.DynamicInvoke(workbook, description, sheet, success);
            }
        }

        private void PublicDefaultHandler(IWorkbook workbook, string desciption, string sheet, bool success)
        {

        }
    }
}
