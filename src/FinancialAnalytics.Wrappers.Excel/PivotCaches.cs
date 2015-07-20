using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotCaches : EntitiesCollectionWrapperBase<IPivotCaches, IPivotCache>, IPivotCaches
    {
        protected ExcelEntityResolver EntityResolver { get; private set; }
        private readonly Microsoft.Office.Interop.Excel.PivotCaches _excelPivotCaches;
        private readonly LateBindingInvoker _invoker;

        public PivotCaches(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotCaches pivotCaches)
            : base()
        {
            if (entityResolver == null)
            {
                throw new ArgumentNullException("entityResolver");
            }
            if (pivotCaches == null)
            {
                throw new ArgumentNullException("pivotCaches");
            }
            this.EntityResolver = entityResolver;
            _excelPivotCaches = pivotCaches;
            _invoker = new LateBindingInvoker(_excelPivotCaches);
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPivotCaches.Count; i ++)
                {
                    IPivotCache pivotCache = EntityResolver.ResolvePivotCache(_excelPivotCaches.Item(i));
                    _items.Add(pivotCache);
                }
            }

        }

        #region EntitiesCollectionWrapperBase members

        public override bool Equals(IPivotCaches obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotCaches pivotCaches = (PivotCaches) obj;
            return _excelPivotCaches.Equals(pivotCaches._excelPivotCaches);
        }

        #endregion

        private bool _disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelPivotCaches);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        public IPivotCache Create(PivotTableSourceType sourceType, Object sourceData, Object version)
        {
            using (new EnUsCultureInvoker())
            {
                IPivotCache pivotCache = EntityResolver.ResolvePivotCache(
                    _invoker.NamedInvoke("Create", XlPivotTableSourceTypeToPivotTableSourceTypeConverter.ConvertBack(sourceType),
                                                            (sourceData is IWorkbookConnection ? ((IWorkbookConnection)sourceData).WorkbookConnectionObject : sourceData),
                                                           version
                                                       ) as Microsoft.Office.Interop.Excel.PivotCache);
                _items.Add(pivotCache);
                return pivotCache;

            }
        }

        public IPivotCache Add(
	        PivotTableSourceType sourceType, 
	        Object sourceData)
        {
            using (new EnUsCultureInvoker())
            {
                IPivotCache pivotCache = EntityResolver.ResolvePivotCache(_excelPivotCaches.Add(XlPivotTableSourceTypeToPivotTableSourceTypeConverter.ConvertBack(sourceType), 
                    (sourceData is IWorkbookConnection  ? ((IWorkbookConnection)sourceData).WorkbookConnectionObject : sourceData)));
                _items.Add(pivotCache);
                return pivotCache;
            }
        }
    }
}
