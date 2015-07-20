using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Sheets : EntitiesCollectionWrapperBase<ISheets, ISheet>, ISheets
    {
        protected ExcelEntityResolver _entityResolver;
        private Microsoft.Office.Interop.Excel.Sheets _excelSheets;

        public Sheets(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Sheets sheets)
        {
            if (sheets == null)
                throw new ArgumentNullException("sheets");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelSheets = sheets;
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
				ComObjectsFinalizer.ReleaseComObject(_excelSheets);
			    _excelSheets = null;
                _entityResolver = null;
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelSheets.Count; i++)
                {
                    ISheet sheet = null;
					dynamic excelSheet = _excelSheets[i];
                    if (excelSheet is Microsoft.Office.Interop.Excel.Worksheet)
                    {
                        sheet =
							_entityResolver.ResolveWorksheet(excelSheet as Microsoft.Office.Interop.Excel.Worksheet);
                    }
					else if (excelSheet is Microsoft.Office.Interop.Excel.Chart)
                    {
						sheet = _entityResolver.ResolveChart(excelSheet as Microsoft.Office.Interop.Excel.Chart);
                    }
					if (sheet != null)
					{
						_items.Add(sheet);
					}
					else
					{
						Marshal.ReleaseComObject(excelSheet);
					}
                }
            }
        }

        public ISheet AddLast()
        {
            using (new EnUsCultureInvoker())
            {
                ISheet sheet =
                    _entityResolver.ResolveWorksheet(
                        _excelSheets.Add(Type.Missing, _excelSheets[_excelSheets.Count], Type.Missing, Type.Missing) as
                        Microsoft.Office.Interop.Excel.Worksheet);
                _items.Add(sheet);
                return sheet;
            }
        }

        public ISheet AddFirst()
        {
            using (new EnUsCultureInvoker())
            {
                ISheet sheet =
                    _entityResolver.ResolveWorksheet(
                        _excelSheets.Add(Type.Missing, _excelSheets[_excelSheets.Count], Type.Missing, Type.Missing) as
                        Microsoft.Office.Interop.Excel.Worksheet);
                _items.Add(sheet);
                return sheet;
            }
        }

        public override bool Equals(ISheets obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Sheets sheets = (Sheets)obj;
            return _excelSheets.Equals(sheets._excelSheets);
        }

        //public ISheet Add(object before, object after)
        //{
        //    ISheet sheet = EntitiesContainer.Resolve<IWorksheet>(new ParameterOverride("worksheet", _excelSheets.Add(before, after)));
        //    Add(sheet);
        //    return sheet;
        //}

		public ISheet this[string codeName]
		{
			get
			{
				Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelSheets.Item[codeName];
				return _entityResolver.ResolveWorksheet(sheet);
			}
		}
	}
}
