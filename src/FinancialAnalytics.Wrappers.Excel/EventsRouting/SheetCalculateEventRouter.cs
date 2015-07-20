using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetCalculateEventHandler(ISheet sheet);

    internal class SheetCalculateEventRouter : ExcelBaseEvent
    {
        public SheetCalculateEventRouter()
            : base(0x0000061B)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetCalculateEventHandler(SheetPrivateHandler);
            PublicEvent = new SheetCalculateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet)
        {
            
        }
    }
}
