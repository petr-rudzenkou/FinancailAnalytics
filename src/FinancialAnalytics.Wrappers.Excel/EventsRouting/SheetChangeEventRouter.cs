using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetChangeEventHandler(ISheet sheet, IRange range);

    internal class SheetChangeEventRouter : ExcelBaseEvent
    {
        public SheetChangeEventRouter()
            : base(0x0000061C)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(SheetRangePrivateHandler);
            PublicEvent = new SheetChangeEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet, IRange range)
        {
            
        }
    }
}
