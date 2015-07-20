using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ErrorBars : ExcelEntityWrapper<IErrorBars>, IErrorBars
    {
        protected Microsoft.Office.Interop.Excel.ErrorBars _excelErrorBars;

        public ErrorBars(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ErrorBars errorBars)
            : base(entityResolver)
        {
            if (errorBars == null)
            {
                throw new ArgumentNullException("errorBars");
            }
            _excelErrorBars = errorBars;
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelErrorBars.Border);
                }
            }
        }

        public override bool Equals(IErrorBars obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ErrorBars bars = (ErrorBars)obj;
            return _excelErrorBars.Equals(bars._excelErrorBars);
        }

        #region Disposable pattern

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelErrorBars);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
