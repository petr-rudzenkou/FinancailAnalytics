using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetPivotTableUpdateEventHandler(ISheet sheet, IPivotTable target);

    internal class SheetPivotTableUpdateEventRouter : ExcelBaseEvent
    {
        public SheetPivotTableUpdateEventRouter()
            : base(0x0000086D)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableUpdateEventHandler(PrivateHandler);
            PublicEvent = new SheetPivotTableUpdateEventHandler(PublicDefaultHandler);
        }

        private void PrivateHandler(object excelSheet, Microsoft.Office.Interop.Excel.PivotTable excelPivotTable)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                ISheet sheet;
                IPivotTable pivotTable;
                lock (Lock)
                {
                    sheet = EntityResolver.ResolveSheet(excelSheet);
                    pivotTable = EntityResolver.ResolvePivotTable(excelPivotTable);
                }
                PublicEvent.DynamicInvoke(sheet, pivotTable);
            }
        }

        private void PublicDefaultHandler(ISheet sheet, IPivotTable target)
        {

        }
    }
}
