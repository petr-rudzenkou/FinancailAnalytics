using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using System.Linq;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Panes : EntitiesCollectionWrapperBase<IPanes, IPane>, IPanes
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Panes _excelPanes;

        public Panes(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Panes panes)
        {
            if (panes == null)
                throw new ArgumentNullException("panes");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelPanes = panes;
            _entityResolver = entityResolver;
            InitializeCollection();
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
                ComObjectsFinalizer.ReleaseComObject(_excelPanes);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion Disposable pattern

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPanes.Count; i++)
                {
                    IPane pane = _entityResolver.ResolvePane(_excelPanes[i]);
                    _items.Add(pane);
                }
            }
        }

        public override bool Equals(IPanes obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Panes panes = (Panes)obj;
            return _excelPanes.Equals(panes._excelPanes);
        }
    }
}
