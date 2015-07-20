using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotItems : EntitiesCollectionWrapperBase<IPivotItems, IPivotItem>, IPivotItems
    {
        public ExcelEntityResolver EntityResolver { get; private set; }
        private readonly MSExcel.PivotItems _excelPivotItems;

        public PivotItems(ExcelEntityResolver entityResolver, MSExcel.PivotItems excelPivotItems) : base()
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            if (excelPivotItems == null)
                throw new ArgumentNullException("excelPivotItems");
            EntityResolver = entityResolver;
            _excelPivotItems = excelPivotItems;
            InitializeCollection();
        }

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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotItems);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPivotItems.Count; i++)
                {
                    AddItemToCollection(_excelPivotItems.Item(i) as MSExcel.PivotItem);
                }
            }
        }

        private void AddItemToCollection(MSExcel.PivotItem pivotItem)
        {
            _items.Add(
                EntityResolver.ResolvePivotItem(pivotItem)
                );
        }

        public override bool Equals(IPivotItems obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotItems pivotItems = (PivotItems)obj;
            return _excelPivotItems.Equals(pivotItems._excelPivotItems);
        }

        public new int Count
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotItems.Count;
                }
            }
        }
    }
}
