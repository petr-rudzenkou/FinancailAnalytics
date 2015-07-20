using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    public abstract class ExcelEntityWrapper<T> : EntityWrapperBase<T>
    {
        public new ExcelEntityResolver EntityResolver
        {
            get { return base._entityResolver as ExcelEntityResolver; }
        }

        protected ExcelEntityWrapper(ExcelEntityResolver entityResolver)
            :base (entityResolver)
        {

        }

		#region Disposable pattern
		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					// Release managed resources.
				}
				// Release unmanaged resources.
				// Set large fields to null.
				// Call Dispose on your base class.
				disposed = true;
			}
			base.Dispose(disposing);
		}
		// The derived class does not have a Finalize method
		// or a Dispose method without parameters because it inherits
		// them from the base class.
		#endregion
	}
}
