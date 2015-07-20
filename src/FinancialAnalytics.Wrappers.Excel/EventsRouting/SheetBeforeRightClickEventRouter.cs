using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetBeforeRightClickEventHandler(ISheet sheet, IRange range, ref bool cancel);

    internal class SheetBeforeRightClickEventRouter : ExcelBaseEvent
    {
        public SheetBeforeRightClickEventRouter()
            : base(0x00000618)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(SheetRangeRefBoolPrivateHandler);
            PublicEvent = new SheetBeforeRightClickEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet, IRange range, ref bool cancel)
        {
            
        }
    }
}
