using System;
using System.Collections;
using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    public abstract class EntitiesCollectionWrapperBase<TCollection, TItem> : IEntitiesCollectionWrapper<TCollection, TItem> where TItem : IEntityWrapper<TItem>
    {
        protected IList<TItem> _items = new List<TItem>();


		#region Disposable pattern
		private bool disposed = false;

		protected virtual void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					// Free other state (managed objects).
				}
				// Free your own state (unmanaged objects).
				// Set large fields to null.
				disposed = true;
			}
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		~EntitiesCollectionWrapperBase()
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

        public void FullDispose()
        {
            foreach (TItem item in _items)
            {
                item.Dispose();
            }
            Dispose();
        }

        public virtual IEnumerator<TItem> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public abstract bool Equals(TCollection obj);

        public virtual TItem this[int index]
        {
            get { return _items[index]; }
        }

        public virtual int Count
        {
            get { return _items.Count; }
        }
    }
}
