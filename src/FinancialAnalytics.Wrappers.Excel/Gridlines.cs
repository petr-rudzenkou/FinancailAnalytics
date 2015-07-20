using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Gridlines : ExcelEntityWrapper<IGridlines>, IGridlines
    {
        private Microsoft.Office.Interop.Excel.Gridlines _excelGridlines;

        public Gridlines(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Gridlines gridlines)
            : base(entityResolver)
        {
            if (gridlines == null)
                throw new ArgumentNullException("gridlines");
            _excelGridlines = gridlines;
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelGridlines.Border);
                }
            }
        }

        public override bool Equals(IGridlines obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Gridlines gridlines = (Gridlines)obj;
            return _excelGridlines.Equals(gridlines._excelGridlines);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelGridlines);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
