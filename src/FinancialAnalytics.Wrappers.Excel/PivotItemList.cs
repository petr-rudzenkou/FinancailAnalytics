using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotItemList : EntitiesCollectionWrapperBase<IPivotItemList, IPivotItem>, IPivotItemList
    {
        private readonly MSExcel.PivotItemList _excelPivotItemList;
        protected ExcelEntityResolver EntityResolver { get; private set; }

        public PivotItemList(ExcelEntityResolver entityResolver, MSExcel.PivotItemList excelPivotItemList)
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            if (excelPivotItemList == null)
                throw new ArgumentNullException("excelPivotItemList");
            EntityResolver = entityResolver;
            _excelPivotItemList = excelPivotItemList;
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPivotItemList.Count; i ++)
                {
                    AddItemToCollection(_excelPivotItemList.Item(i));
                }
            }
        }

        private void AddItemToCollection(MSExcel.PivotItem pivotItem)
        {
            _items.Add(
                EntityResolver.ResolvePivotItem(pivotItem)
                );
        }

        public override bool Equals(IPivotItemList obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotItemList pivotItemList = (PivotItemList)obj;
            return _excelPivotItemList.Equals(pivotItemList._excelPivotItemList);
        }

        public new int Count
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotItemList.Count;
                }
            }
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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotItemList);
                _disposed = true;
            }
            base.Dispose(disposing);
        }


        #region IEnumerable Members

        public new System.Collections.IEnumerator GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion
    }
}
