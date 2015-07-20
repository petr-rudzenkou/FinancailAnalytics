using System;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.EventsRouting
{
    internal abstract class ExcelBaseEvent : BaseEvent
    {
        private static readonly WorkbooksCollector WorkbooksCollector;
        private static readonly Guid EventsUid;

        static ExcelBaseEvent()
        {
            WorkbooksCollector = new WorkbooksCollector();
            EventsUid = new Guid(Application.EXCEL_APPLICATION_EVENTS_INTERFACE_GUID);
        }

        protected ExcelBaseEvent(int dispId)
            : base(dispId, EventsUid)
        {
        }

        protected void WorkbookPrivateHandler(Microsoft.Office.Interop.Excel.Workbook excelWorkbook)
        {
            WorkbooksCollector.Stop();
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook;
                lock (Lock)
                {
                    workbook = EntityResolver.ResolveWorkbook(excelWorkbook);
                    WorkbooksCollector.Add(workbook);
                }
                PublicEvent.DynamicInvoke(workbook);
                WorkbooksCollector.Start();
            }
            
        }

        protected void WorkbookRefBoolPrivateHandler(Microsoft.Office.Interop.Excel.Workbook excelWorkbook, ref bool cancel)
        {
            WorkbooksCollector.Stop();
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook;
                lock (Lock)
                {
                    workbook = EntityResolver.ResolveWorkbook(excelWorkbook);
                    WorkbooksCollector.Add(workbook);
                }
                object[] args = { workbook, cancel };
                PublicEvent.DynamicInvoke(args);
                cancel = (bool)args[1];
                WorkbooksCollector.Start();
            }
        }

        protected void WorkbookWindowPrivateHandler(Microsoft.Office.Interop.Excel._Workbook excelWorkbook, Microsoft.Office.Interop.Excel.Window excelWindow)
        {
            WorkbooksCollector.Stop();
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook;
                IWindow window;
                lock (Lock)
                {
                    workbook = EntityResolver.ResolveWorkbook(excelWorkbook);
                    WorkbooksCollector.Add(workbook);
                    window = EntityResolver.ResolveWindow(excelWindow);
                }
                PublicEvent.DynamicInvoke(workbook, window);
                WorkbooksCollector.Start();
            }
        }

        protected void SheetPrivateHandler(object excelSheet)
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
                PublicEvent.DynamicInvoke(sheet);
            }
        }

        protected void SheetRangePrivateHandler(object excelSheet, Microsoft.Office.Interop.Excel.Range excelRange)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                ISheet sheet;
                IRange range;
                lock (Lock)
                {
                    sheet = EntityResolver.ResolveSheet(excelSheet);
                    range = EntityResolver.ResolveRange(excelRange);
                }
                PublicEvent.DynamicInvoke(sheet, range);
            }
        }

        protected void SheetRangeRefBoolPrivateHandler(object excelSheet, Microsoft.Office.Interop.Excel.Range excelRange, ref bool cancel)
        {
            if (!IsEnabled)
            {
                return;
            }
            using (new EnUsCultureInvoker())
            {
                ISheet sheet;
                IRange range;
                lock (Lock)
                {
                    sheet = EntityResolver.ResolveSheet(excelSheet);
                    range = EntityResolver.ResolveRange(excelRange);
                }
                object[] args = { sheet, range, cancel };
                PublicEvent.DynamicInvoke(args);
                cancel = (bool)args[2];
            }
        }


    }
}
