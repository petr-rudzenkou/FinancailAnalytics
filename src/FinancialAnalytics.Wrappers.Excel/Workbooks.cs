using System;
using System.Globalization;
using System.Linq;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Workbooks : EntitiesCollectionWrapperBase<IWorkbooks, IWorkbook>, IWorkbooks
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Workbooks _excelWorkbooks;
        private static readonly Object _locker = new object();

        public Workbooks(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Workbooks workbooks)
        {
            if (workbooks == null)
                throw new ArgumentNullException("workbooks");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelWorkbooks = workbooks;
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
				ComObjectsFinalizer.ReleaseComObject(_excelWorkbooks);
			    _excelWorkbooks = null;
                _entityResolver = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public IWorkbook Open(string fileName)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel._Workbook excelWorkbook = _excelWorkbooks.Open(fileName, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                IWorkbook workbook = _entityResolver.ResolveWorkbook(excelWorkbook);
                _items.Add(workbook);
                return workbook;
            }
        }

        public IWorkbook Add()
        {
            using (new EnUsCultureInvoker())
            {
                //Word "Workbook" need to be translated into Office UI language, otherwise there is an exception.
                //We need to pass word "Workbook" as a parameter to support proper embedding chart colors in Excel and PP
                CultureInfo culture = _entityResolver.ResolveApplication().UICulture;
                object templateName = WorkbookTemplateFactory.GetDefaultTemplate(culture);
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = _excelWorkbooks.Add(templateName);
                IWorkbook workbook = _entityResolver.ResolveWorkbook(excelWorkbook);
                _items.Add(workbook);
                return workbook;
            }
        }

        public IWorkbook Add(string templateName)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
                lock(_locker)
                {
                    excelWorkbook = _excelWorkbooks.Add(templateName);
                }
                IWorkbook workbook = _entityResolver.ResolveWorkbook(excelWorkbook);
                _items.Add(workbook);
                return workbook;
            }
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelWorkbooks.Count; i++)
                {
                    IWorkbook workbook = _entityResolver.ResolveWorkbook(_excelWorkbooks[i]);
                    _items.Add(workbook);
                }
            }
        }

        public IWorkbook this[string fullName]
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return GetWorkbookByFullName(fullName);
                }
            }
        }

        public bool Contains(string fullName)
        {
            using (new EnUsCultureInvoker())
            {
                IWorkbook workbook = GetWorkbookByFullName(fullName);
                return workbook != null;
            }
        }

        protected virtual IWorkbook GetWorkbookByFullName(string workbookFullName)
        {
            using (new EnUsCultureInvoker())
            {
                return (from workbook in this
						where workbook.UncFullName.Equals(workbookFullName, StringComparison.InvariantCultureIgnoreCase)
                        select workbook).FirstOrDefault();
            }
        }

        public override bool Equals(IWorkbooks obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Workbooks chartTitle = (Workbooks)obj;
            return _excelWorkbooks.Equals(chartTitle._excelWorkbooks);
        }

        public void CloseAll()
        {
            foreach (IWorkbook workbook in this)
            {
                workbook.Close(false);
            }
        }

		public void CloseAllWithoutAlerts()
		{
			foreach (IWorkbook workbook in this)
			{
				workbook.CloseWithoutAlerts(false);
			}
		}
    }
}
