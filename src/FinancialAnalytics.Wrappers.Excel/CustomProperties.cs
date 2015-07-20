using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class CustomProperties : EntitiesCollectionWrapperBase<ICustomProperties, ICustomProperty>, ICustomProperties
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.CustomProperties _excelCustomProperties;

        public CustomProperties(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.CustomProperties customProperties)
        {
            if (customProperties == null)
                throw new ArgumentNullException("customProperties");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelCustomProperties = customProperties;
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
				ComObjectsFinalizer.ReleaseComObject(_excelCustomProperties);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelCustomProperties.Count; i++)
                {
                    ICustomProperty customProperty = _entityResolver.ResolveCustomProperty(_excelCustomProperties[i]);
                    _items.Add(customProperty);
                }
            }
        }

        public ICustomProperty Add(string name, string value)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.CustomProperty excelCustomProperty = _excelCustomProperties.Add(name,
                                                                                                               value);
                ICustomProperty customProperty = _entityResolver.ResolveCustomProperty(excelCustomProperty);
                _items.Add(customProperty);
                return customProperty;
            }
        }

        public override bool Equals(ICustomProperties obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CustomProperties application = (CustomProperties)obj;
            return _excelCustomProperties.Equals(application._excelCustomProperties);
        }
    }
}
