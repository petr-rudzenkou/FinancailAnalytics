using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="saveUI">True if the Save As dialog box will be displayed.</param>
    /// <param name="cancel">False when the event occurs. If the event procedure sets this argument to True, the workbook isn't saved when the procedure is finished.</param>
    public delegate void WorkbookBeforeSaveEventHandler(IWorkbook workbook, bool saveUi, ref bool cancel);

    internal class WorkbookBeforeSaveEventRouter : ExcelBaseEvent
    {
        public WorkbookBeforeSaveEventRouter()
            : base(0x00000623)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(PrivateHandler);
            PublicEvent = new WorkbookBeforeSaveEventHandler(PublicDefaultHandler);
        }

        private void PrivateHandler(Microsoft.Office.Interop.Excel.Workbook excelWorkbook, bool saveUi, ref bool cancel)
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
                object[] args = { workbook, saveUi, cancel };
                PublicEvent.DynamicInvoke(args);
                cancel = (bool)args[2];
            }
        }

        private void PublicDefaultHandler(IWorkbook workbook, bool saveUi, ref bool cancel)
        {

        }
    }
}
