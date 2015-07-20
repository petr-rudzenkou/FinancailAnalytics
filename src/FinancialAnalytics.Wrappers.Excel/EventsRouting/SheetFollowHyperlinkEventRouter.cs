using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    public delegate void SheetFollowHyperlinkEventHandler(ISheet sheet, object target);

    internal class SheetFollowHyperlinkEventRouter : ExcelBaseEvent
    {
        public SheetFollowHyperlinkEventRouter()
            : base(0x0000073E)
        {
            PrivateEvent = new Microsoft.Office.Interop.Excel.AppEvents_SheetFollowHyperlinkEventHandler(PrivateHandler);
            PublicEvent = new SheetFollowHyperlinkEventHandler(PublicDefaultHandler);
        }

        private void PrivateHandler(object excelSheet, Microsoft.Office.Interop.Excel.Hyperlink excelHyperlink)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                ISheet sheet;
                lock (Lock)
                {
                    sheet = EntityResolver.ResolveSheet(excelSheet);
                }
                PublicEvent.DynamicInvoke(sheet, excelHyperlink);
            }
        }

        private void PublicDefaultHandler(ISheet sheet, object target)
        {

        }
    }
}
