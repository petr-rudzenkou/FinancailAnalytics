using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    /// <summary>
    /// The sheet. Can be a Chart or Worksheet object.
    /// </summary>
    /// <param name="sheet"></param>
    public delegate void SheetDeactivateEventHandler(ISheet sheet);

    internal class SheetDeactivateEventRouter : ExcelBaseEvent
    {
        public SheetDeactivateEventRouter()
            : base(0x0000061A)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetDeactivateEventHandler(SheetPrivateHandler);
            PublicEvent = new SheetDeactivateEventHandler(PublicDefaultHandler);
        }

        private void PublicDefaultHandler(ISheet sheet)
        {
            
        }
    }
}
