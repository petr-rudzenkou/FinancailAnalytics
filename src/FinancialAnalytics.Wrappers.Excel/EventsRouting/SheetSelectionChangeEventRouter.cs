using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetSelectionChangeEventHandler(ISheet sheet, IRange range);

    internal class SheetSelectionChangeEventRouter : ExcelBaseEvent
    {
        public SheetSelectionChangeEventRouter()
            : base(0x00000616)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(SheetRangePrivateHandler);
            PublicEvent = new SheetSelectionChangeEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet, IRange range)
        {
            
        }
    }
}
