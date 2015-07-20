using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class Worksheets : EntitiesCollectionWrapperBase<IWorksheets, IWorksheet>, IWorksheets
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Sheets _excelSheets;

        public Worksheets(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Sheets sheets)
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

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
			{
				for (int i = 1; i <= _excelSheets.Count; i++)
				{
					dynamic sheet = _excelSheets[i];
					if (sheet is Microsoft.Office.Interop.Excel.Worksheet)
				    {
						IWorksheet worksheet =
							_entityResolver.ResolveWorksheet(sheet as Microsoft.Office.Interop.Excel.Worksheet);
						_items.Add(worksheet);
					}
					else
					{
						Marshal.ReleaseComObject(sheet);
					}
				}
            }
        }

        public IWorksheet this[string codeName]
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _entityResolver.ResolveWorksheet(_excelSheets[codeName] as Microsoft.Office.Interop.Excel.Worksheet);
                }
            }
        }

        public IWorksheet AddLast()
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Worksheet newExcelWorksheet;
                if (_excelSheets.Count == 0)
                {
                    newExcelWorksheet =
                        (Microsoft.Office.Interop.Excel.Worksheet) _excelSheets.Add(Type.Missing, Type.Missing,
                        Type.Missing, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                }
                else
                {
                    try
                    {
                        newExcelWorksheet =
                        (Microsoft.Office.Interop.Excel.Worksheet)
                        _excelSheets.Add(Type.Missing, _excelSheets[_excelSheets.Count],
                                         Type.Missing, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                IWorksheet worksheet = _entityResolver.ResolveWorksheet(newExcelWorksheet);
                _items.Add(worksheet);
                return worksheet;
            }
        }

        public IWorksheet AddFirst()
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Worksheet newExcelWorksheet;
                if (_excelSheets.Count == 0)
                {
                    newExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelSheets.Add(System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, 
                        Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                }
                else
                {
                    newExcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelSheets.Add(1,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value, 
                        Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                }
                IWorksheet worksheet = _entityResolver.ResolveWorksheet(newExcelWorksheet);
                _items.Add(worksheet);
                return worksheet;
            }
        }

        public Object Add(Object before, Object after, Object count, Object type)
        {
            using (new EnUsCultureInvoker())
            {
                IWorksheet worksheet = null;
                if ((type as string) == null)
                {
                     worksheet = _entityResolver.ResolveWorksheet(_excelSheets.Add(before, after, count, type) as Microsoft.Office.Interop.Excel.Worksheet);
                }
                else
                {
                    //case we add a template: add at end
                    var sheets = _excelSheets.Application.Sheets;
                    worksheet = _entityResolver.ResolveWorksheet(sheets.Add(before, (sheets.Count == 0) ? Type.Missing : sheets[sheets.Count], count, type) as Microsoft.Office.Interop.Excel.Worksheet);
                }
                _items.Add(worksheet);
                return worksheet;
            }
        }

        public IWorksheet Add()
        {
            using (new EnUsCultureInvoker())
            {
                IWorksheet worksheet = _entityResolver.ResolveWorksheet(_excelSheets.Add() as Microsoft.Office.Interop.Excel.Worksheet);
                _items.Add(worksheet);
                return worksheet;
            }
        }

        public override bool Equals(IWorksheets obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Worksheets worksheets = (Worksheets)obj;
            return _excelSheets.Equals(worksheets._excelSheets);
        }
    }
}
