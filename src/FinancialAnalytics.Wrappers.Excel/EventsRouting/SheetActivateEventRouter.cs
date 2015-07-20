using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetActivateEventHandler(ISheet sheet);

    internal class SheetActivateEventRouter : ExcelBaseEvent
    {
        public SheetActivateEventRouter()
            : base(0x00000619)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(SheetPrivateHandler);
            PublicEvent = new SheetActivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet)
        {
            
        }
    }
}
