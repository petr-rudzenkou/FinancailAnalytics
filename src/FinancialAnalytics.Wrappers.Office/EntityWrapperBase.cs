using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
    [ComVisible(true)]
    public abstract class EntityWrapperBase<T> : IEntityWrapper<T>
    {
        protected EntityResolverBase _entityResolver;

        public EntityResolverBase EntityResolver
        {
            get { return _entityResolver; }
        }

        protected EntityWrapperBase(EntityResolverBase entityResolver)
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _entityResolver = entityResolver;
        }

		#region Disposable pattern
		private bool disposed = false;

		protected virtual void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					_entityResolver = null;
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

		~EntityWrapperBase()
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

		
        public abstract bool Equals(T obj);
    }
}
