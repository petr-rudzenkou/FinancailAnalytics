using System;
using System.Collections;
using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    public abstract class LazyEntitiesCollectionWrapper<TCollection, TItem> : IEntitiesCollectionWrapper<TCollection, TItem> where TItem : IEntityWrapper<TItem>
    {
        protected IList<TItem> _items;
        protected bool _initialized;

        public IEnumerator<TItem> GetEnumerator()
        {
            if (!_initialized)
                InitializeCollection();

            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            if (!_initialized)
                InitializeCollection();

            return _items.GetEnumerator();
        }

        public virtual TItem this[int index]
        {
            get
            {
                if (!_initialized)
                    InitializeCollection();

                return _items[index];
            }
        }

        public virtual int Count
        {
            get
            {
                if (!_initialized)
                    InitializeCollection();

                return _items.Count;
            }
        }

        public abstract bool Equals(TCollection obj);

        protected virtual void InitializeCollection()
        {
            _items = new List<TItem>();
            _initialized = true;
        }

        public void FullDispose()
        {
            if (!_initialized)
                return;

            foreach (TItem item in _items)
            {
                item.Dispose();
            }
            Dispose();
        }

        #region Disposable pattern

        private bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Free other state (managed objects).
                }
                // Free your own state (unmanaged objects).
                // Set large fields to null.
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~LazyEntitiesCollectionWrapper()
        {
            try
            {
                Dispose(false);
            }
            catch (Exception)
            {
            }
        }

        #endregion
    }
}
