using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ListObjects : EntitiesCollectionWrapperBase<IListObjects, IListObject>, IListObjects
    {
        protected ExcelEntityResolver _entityResolver;

        protected Microsoft.Office.Interop.Excel.ListObjects _excelListObjects;

        public ListObjects(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ListObjects listObjects)
        {
            if (listObjects == null)
            {
                throw new ArgumentNullException("listObjects");
            }
            if (entityResolver == null)
            {
                throw new ArgumentNullException("entityResolver");
            }
            _excelListObjects = listObjects;
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
                ComObjectsFinalizer.ReleaseComObject(_excelListObjects);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelListObjects.Count; i++)
                {
                    IListObject listObject = _entityResolver.ResolveListObject(_excelListObjects[i]);
                    _items.Add(listObject);
                }
            }
        }

        public override bool Equals(IListObjects obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ListObjects windows = (ListObjects)obj;
            return _excelListObjects.Equals(windows._excelListObjects);
        }
    }
}