using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class OLEObjects : LazyEntitiesCollectionWrapper<IOLEObjects, IOLEObject>, IOLEObjects 
    {
        protected ExcelEntityResolver _entityResolver;
        Microsoft.Office.Interop.Excel.OLEObjects _excelOLEObjects;

        public OLEObjects(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.OLEObjects oleObjects)
        {
            if (oleObjects == null)
                throw new ArgumentNullException("oleObjects");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelOLEObjects = oleObjects;
            _entityResolver = entityResolver;
        }

        public IOLEObject Add()
        {
            using (new EnUsCultureInvoker())
            {
                IOLEObject oleObj = _entityResolver.ResolveOLEObject(_excelOLEObjects.Add (Type.Missing, Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing) as Microsoft.Office.Interop.Excel.OLEObject);
                _items.Add(oleObj);
                return oleObj;
            }
        }

        public override int Count
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelOLEObjects.Count;
                }
            }
        }

        protected override void InitializeCollection()
        {
            base.InitializeCollection();
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelOLEObjects.Count; i++)
                {
					IOLEObject shape = _entityResolver.ResolveOLEObject(_excelOLEObjects.Item(i) as Microsoft.Office.Interop.Excel.OLEObject);
                    _items.Add(shape);
                }
            }
        }

        public override bool Equals(IOLEObjects obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            OLEObjects charts = (OLEObjects)obj;
            return _excelOLEObjects.Equals(charts._excelOLEObjects);
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
                ComObjectsFinalizer.ReleaseComObject(_excelOLEObjects);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
