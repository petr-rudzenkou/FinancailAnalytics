using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetBeforeDoubleClickEventHandler(ISheet sheet, IRange range, ref bool cancel);

    internal class SheetBeforeDoubleClickEventRouter : ExcelBaseEvent
    {
        public SheetBeforeDoubleClickEventRouter()
            : base(0x00000617)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeDoubleClickEventHandler(SheetRangeRefBoolPrivateHandler);
            PublicEvent = new SheetBeforeDoubleClickEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet, IRange range, ref bool cancel)
        {
            
        }
    }
}
