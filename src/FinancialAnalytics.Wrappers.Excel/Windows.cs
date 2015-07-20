using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using System.Linq;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Windows : EntitiesCollectionWrapperBase<IWindows, IWindow>, IWindows
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Windows _excelWindows;

        public Windows(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Windows windows)
        {
            if (windows == null)
                throw new ArgumentNullException("windows");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelWindows = windows;
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
				ComObjectsFinalizer.ReleaseComObject(_excelWindows);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelWindows.Count; i++)
                {
                    IWindow window = _entityResolver.ResolveWindow(_excelWindows[i]);
                    _items.Add(window);
                }
            }
        }

		public override bool Equals(IWindows obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Windows windows = (Windows)obj;
            return _excelWindows.Equals(windows._excelWindows);
        }

        public IWindow this[String name]
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                        return
                        _items.SingleOrDefault(item => item.Caption == name);
                }
            }
        }
    }
}
