using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ListObject : ExcelEntityWrapper<IListObject>, IListObject
    {
        protected Microsoft.Office.Interop.Excel.ListObject _excelListObject;

        public ListObject(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ListObject listObject)
            :base(entityResolver)
        {
            if (listObject == null)
                throw new ArgumentNullException("listObject");
            _excelListObject = listObject;
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
				ComObjectsFinalizer.ReleaseComObject(_excelListObject);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public IRange Range
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelListObject.Range);
                }
            }
        }

        public IRange DataBodyRange
        {
            get 
            {
                using (new EnUsCultureInvoker())
                {
                    try
                    {
                        // Ex2003 throw exception when ListObject consist of one empty cell
                        Microsoft.Office.Interop.Excel.Range dataBodyRange = _excelListObject.DataBodyRange;
                        if (dataBodyRange != null)
                        {
                            return EntityResolver.ResolveRange(dataBodyRange);
                        }
                    }
                    catch (COMException ex)
                    {
                        bool rethrow = ExceptionHandler.HandleException(ex);
                        if (rethrow)
                            throw;
                    }
                    return null;
                }
            }
        }

        public IRange HeaderRowRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelListObject.HeaderRowRange);
                }
            }
        }

        public IListRows ListRows
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveListRows(_excelListObject.ListRows);
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelListObject.Name;
                }
            }
        }

        public bool Active
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelListObject.Active;
                }
            }
        }
        
     
        public override bool Equals(IListObject obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ListObject chartTitle = (ListObject)obj;
            return _excelListObject.Equals(chartTitle._excelListObject);
        }
    }
}
