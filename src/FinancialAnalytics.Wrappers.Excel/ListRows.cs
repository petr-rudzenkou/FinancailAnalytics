using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class ListRows : ExcelEntityWrapper<IListRows>, IListRows
    {
        private Microsoft.Office.Interop.Excel.ListRows _excelListRows;

        public ListRows(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ListRows listRows)
            : base(entityResolver)
        {
            if (listRows == null)
            {
                throw new ArgumentNullException("listRows");
            }
            _excelListRows = listRows;
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelListRows);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion

        public int Count
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelListRows.Count;
                }
            }
        }

        public override bool Equals(IListRows obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ListRows chartTitle = (ListRows)obj;
            return _excelListRows.Equals(chartTitle._excelListRows);
        }
    }
}