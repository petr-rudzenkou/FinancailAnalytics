using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class WorkbookConnection : ExcelEntityWrapper<IWorkbookConnection>, IWorkbookConnection
    {

        private Object _excelWorkbookConnection;
        private LateBindingInvoker _invoker;
        public WorkbookConnection(ExcelEntityResolver entityResolver, Object excelWorkbookConnection)
            : base(entityResolver)
        {
            if (excelWorkbookConnection == null)
            {
                throw new ArgumentNullException("excelWorkbookConnection");
            }
            _excelWorkbookConnection = excelWorkbookConnection;
            _invoker = new LateBindingInvoker(_excelWorkbookConnection);
        }

        public override bool Equals(IWorkbookConnection obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            WorkbookConnection connections = (WorkbookConnection)obj;
            return _excelWorkbookConnection.Equals(connections._excelWorkbookConnection);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelWorkbookConnection);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _invoker.InvokeGetPropertyValue<String>("Name");
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _invoker.InvokeSetPropertyValue("Name", value);
                }
            }
        }

        public Object WorkbookConnectionObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorkbookConnection;
                }
            }
        }

    }
}
