using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class CustomProperty : ExcelEntityWrapper<ICustomProperty>, ICustomProperty
    {
        protected Microsoft.Office.Interop.Excel.CustomProperty _excelCustomProperty;

        public CustomProperty(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.CustomProperty customProperty)
            : base(entityResolver)
        {
            if (customProperty == null)
                throw new ArgumentNullException("customProperty");
            _excelCustomProperty = customProperty;
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
				ComObjectsFinalizer.ReleaseComObject(_excelCustomProperty);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelCustomProperty.Name;
                }
            }
        }

        public string Value
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelCustomProperty.Value.ToString();
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelCustomProperty.Value = value;
                }
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelCustomProperty.Delete();
            }
        }
        
        public override bool Equals(ICustomProperty obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CustomProperty application = (CustomProperty)obj;
            return _excelCustomProperty.Equals(application._excelCustomProperty);
        }
    }
}
