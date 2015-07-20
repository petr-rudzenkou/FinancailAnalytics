using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Names : EntitiesCollectionWrapperBase<INames, IName>, INames
    {
        #region Constants and Fields

        protected ExcelEntityResolver _entityResolver;

        private readonly Microsoft.Office.Interop.Excel.Names _excelNames;

        private bool _disposed;
		
		private bool _isCollectionInitialized;

        #endregion

        #region Constructors and Destructors

        public Names(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Names names)
        {
            if (names == null)
            {
                throw new ArgumentNullException("names");
            }
            if (entityResolver == null)
            {
                throw new ArgumentNullException("entityResolver");
            }
            _excelNames = names;
            _entityResolver = entityResolver;
        }

        #endregion

        #region Implemented Interfaces

        #region IEquatable<INames>

        public override bool Equals(INames obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            var workbook = (Names)obj;
            return _excelNames.Equals(workbook._excelNames);
        }

        #endregion

        #region INames
		
		public override int Count
        {
            get
            {
                if (_isCollectionInitialized)
                {
                    return _items.Count;
                }
                using (new EnUsCultureInvoker())
                {
                    return _excelNames.Count;
                }
            }
        }

        public override IName this[int index]
        {
            get
            {
                if (_isCollectionInitialized)
                {
                    return _items[index];
                }
                using (new EnUsCultureInvoker())
                {
                    //Description: here we use (index + 1) because indexing of Office collections starts with 1.
                    return _entityResolver.ResolveName(_excelNames.Item(index + 1));
                }
            }
        }

        public override System.Collections.Generic.IEnumerator<IName> GetEnumerator()
        {
            if (!_isCollectionInitialized)
            {
                InitializeCollection();
            }
            return _items.GetEnumerator();
        }

        public IName Add(string rangeName, string refersTo, bool visible, string rangeAddress)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Name excelName = _excelNames.Add(
                    rangeName,
                    refersTo,
                    visible,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    rangeAddress,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);
                IName wrappedName = _entityResolver.ResolveName(excelName);
                if (_isCollectionInitialized)
                {
                    _items.Add(wrappedName);
                }
                return wrappedName;
            }
        }

        public IName Add(
            Object name,
            Object refersTo,
            Object visible,
            Object macroType,
            Object shortcutKey,
            Object category,
            Object nameLocal,
            Object refersToLocal,
            Object categoryLocal,
            Object refersToR1C1,
            Object refersToR1C1Local)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Name excelName = _excelNames.Add(
                    name,
                    refersTo,
                    visible,
                    macroType,
                    shortcutKey,
                    category,
                    nameLocal,
                    refersToLocal,
                    categoryLocal,
                    refersToR1C1,
                    refersToR1C1Local
                    );
                IName wrappedName = _entityResolver.ResolveName(excelName);
                if (_isCollectionInitialized)
                {
                    _items.Add(wrappedName);
                }
                return wrappedName;
            }

        }

        #endregion

        #endregion

        #region Methods

		public IName GetItem(object index, object indexLocal, object refersTo)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Name excelName = _excelNames.Item(index, indexLocal, refersTo);
                return _entityResolver.ResolveName(excelName);
            }
        }
		
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelNames);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                foreach (object name in _excelNames)
                {
                    AddItemToCollection(name as Microsoft.Office.Interop.Excel.Name);
                }
				_isCollectionInitialized = true;
            }
        }

        private void AddItemToCollection(Microsoft.Office.Interop.Excel.Name item)
        {
            IName name = _entityResolver.ResolveName(item);
            _items.Add(name);
        }

        #endregion
    }
}