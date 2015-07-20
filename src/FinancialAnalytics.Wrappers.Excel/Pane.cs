using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Pane : ExcelEntityWrapper<IPane>, IPane
    {
        private Microsoft.Office.Interop.Excel.Pane _excelPane;

        public Pane(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Pane pane)
            : base(entityResolver)
        {
            if (pane == null)
                throw new ArgumentNullException("pane");
            _excelPane = pane;
        }

        #region Disposable pattern

        private bool disposed;
        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }

                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelPane);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Overrides

        public override bool Equals(IPane obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;

            Pane pane = (Pane)obj;
            return _excelPane.Equals(pane._excelPane);
        }

        #endregion Overrides

        public int Index
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPane.Index;
                }
            }
        }

        public object PaneObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPane;
                }
            }
        }

        public IRange VisibleRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelPane.VisibleRange);
                }
            }
        }

        public int ScrollColumn
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPane.ScrollColumn;
                }
            }
        }

        public int ScrollRow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPane.ScrollRow;
                }
            }
        }

        public int PointsToScreenPixelsX(int Points)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelPane.PointsToScreenPixelsX(Points);
            }
        }

        public int PointsToScreenPixelsY(int Points)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelPane.PointsToScreenPixelsY(Points);
            }
        }
    }
}
