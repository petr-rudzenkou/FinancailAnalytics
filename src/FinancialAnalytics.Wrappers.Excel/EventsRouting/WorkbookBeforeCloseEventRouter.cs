using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="workbook">The workbook that's being closed</param>
    /// <param name="cancel">False when the event occurs. If the event procedure sets this argument to True, the workbook doesn't close when the procedure is finished.</param>
    public delegate void WorkbookBeforeCloseEventHandler(IWorkbook workbook, ref bool cancel);

    internal class WorkbookBeforeCloseEventRouter : ExcelBaseEvent
    {
        public WorkbookBeforeCloseEventRouter()
            : base(0x00000622)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(WorkbookRefBoolPrivateHandler);
            PublicEvent = new WorkbookBeforeCloseEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(IWorkbook workbook, ref bool cancel)
        {

        }
    }
}
