using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;
using IChartObjects = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartObjects;
using ICustomProperties = FinancialAnalytics.Wrappers.Excel.Interfaces.ICustomProperties;
using IListObjects = FinancialAnalytics.Wrappers.Excel.Interfaces.IListObjects;
using INames = FinancialAnalytics.Wrappers.Excel.Interfaces.INames;
using IPageSetup = FinancialAnalytics.Wrappers.Excel.Interfaces.IPageSetup;
using IRange = FinancialAnalytics.Wrappers.Excel.Interfaces.IRange;
using IShapes = FinancialAnalytics.Wrappers.Excel.Interfaces.IShapes;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class Worksheet : ExcelEntityWrapper<IWorksheet>, IWorksheet
    {
        protected Microsoft.Office.Interop.Excel.Worksheet _excelWorksheet;
        private static readonly object _locker = new object();
        private static int _counter = 0;
        

        public Worksheet(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Worksheet worksheet)
            : base(entityResolver)
        {
            if (worksheet == null)
                throw new ArgumentNullException("worksheet");
            _excelWorksheet = worksheet;
            InstanceNumber = (++_counter);
        }


        protected int InstanceNumber { get; private set; }
        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {                
                //Here we must dispose unmanaged resources and LOH objects
                lock (_locker)
                {
                    ComObjectsFinalizer.ReleaseComObject(_excelWorksheet);
					_excelWorksheet = null;
                    _entityResolver = null;
                }
                disposed = true;
            }
            base.Dispose(disposing);
            
        }

        ~Worksheet()
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

        public object WorksheetObject
        {
            get { return _excelWorksheet; }
        }

        public IChartObjects ChartObjects
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return
                        EntityResolver.ResolveChartObjects(
                            _excelWorksheet.ChartObjects(Type.Missing) as Microsoft.Office.Interop.Excel.ChartObjects);
                }
            }
        }

       

        public IApplication Application
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveApplication();
                }
            }
        }

        public IRange UsedRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    lock (_locker)
                    {
                        return EntityResolver.ResolveRange(_excelWorksheet.UsedRange);
                    }
                }
            }
        }

		public IRange Rows
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					lock (_locker)
					{
						return EntityResolver.ResolveRange(_excelWorksheet.Rows);
					}
				}
			}
		}

		public IRange Columns
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					lock (_locker)
					{
						return EntityResolver.ResolveRange(_excelWorksheet.Columns);
					}
				}
			}
		}

        public IRange Cells
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelWorksheet.Cells);
                }
            }
        }

        public IShapes Shapes
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapes(_excelWorksheet.Shapes);
                }
            }
        }

        public IOLEObjects OLEObjects
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
					return EntityResolver.ResolveOLEObjects(_excelWorksheet.OLEObjects(Type.Missing) as Microsoft.Office.Interop.Excel.OLEObjects);
                }
            }
        }

        public object SheetObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorksheet;
                }
            }
        }

        public void Paste()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Paste(Type.Missing, Type.Missing);
            }
        }

        public void Paste(object range, object value)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Paste(range, value);
            }
        }

		//Changed to accept object parameter as it was incorrect to accept PasteToChartType enum
		public void PasteSpecial(object format)
		{
            using (new EnUsCultureInvoker())
            {
				_excelWorksheet.PasteSpecial(format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
		}

        public IRange GetRange(IRange cell1, IRange cell2)
        {
            using (new EnUsCultureInvoker())
            {
                IRange range =
                EntityResolver.ResolveRange(
                _excelWorksheet.get_Range(cell1.RangeObject, cell2.RangeObject)
                );
                return range;
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
					string result = null;
					int tries = 5;
					while (tries > 0)
					{
						tries--;
						try
						{
							result = _excelWorksheet.Name; //sometimes we are getting exceptions here
							tries = 0;
						}
						catch (Exception ex)
						{
							if (tries == 0) throw ex;
						}
					}
                    return result;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    //lock(_locker)
                    //{
                        _excelWorksheet.Name = value;
                    //}
                }
            }
        }

        public SheetVisibility Visible
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlSheetVisibilityToSheetVisibilityConverter.Convert(_excelWorksheet.Visible);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    lock (_locker)
                    {
                        _excelWorksheet.Visible = XlSheetVisibilityToSheetVisibilityConverter.ConvertBack(value);
                    }
                }
            }
        }

        public IWorkbook Workbook
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return
                        EntityResolver.ResolveWorkbook(_excelWorksheet.Parent as Microsoft.Office.Interop.Excel.Workbook);
                }
            }
        }

        public INames Names
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveNames(_excelWorksheet.Names);
                }
            }
        }

        public IListObjects ListObjects
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveListObjects(_excelWorksheet.ListObjects);
                }
            }
        }

        public ICustomProperties CustomProperties
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveCustomProperties(_excelWorksheet.CustomProperties);
                }
            }
        }

        public string CodeName
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    //if sheetcode is empty
                    if (string.IsNullOrEmpty(_excelWorksheet.CodeName))
                    {
                        //we need to access sheetcode via VBComponents and Excel will generate it automatically
                        var workbook = _excelWorksheet.Parent as Microsoft.Office.Interop.Excel.Workbook;
                        if (workbook != null)
                        {
                            try
                            {
                                return workbook.VBProject.VBComponents.Item(Name).Name;
                            }
                            catch (Exception exc)
                            {
                                bool rethrow = ExceptionHandler.HandleException(exc);
                                if (rethrow)
                                    throw;
                            }
                            finally
                            {
                                ComObjectsFinalizer.ReleaseComObject(workbook);
                            }
                        }
                    }
                    return _excelWorksheet.CodeName;
                }
            }
        }

        public bool ProtectContents
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWorksheet.ProtectContents;
                }
            }
        }

        public IPageSetup PageSetup
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolvePageSetup(_excelWorksheet.PageSetup);
                }
            }
        }

        public virtual void Protect(Object password,
            bool drawingObjects,
            bool contents,
            bool scenarios,
            bool userInterfaceOnly,
            bool allowFormattingCells,
            bool allowFormattingColumns,
            bool allowFormattingRows,
            bool allowInsertingColumns,
            bool allowInsertingRows,
            bool allowInsertingHyperlinks,
            bool allowDeletingColumns,
            bool allowDeletingRows,
            bool allowSorting,
            bool allowFiltering,
            bool allowUsingPivotTables)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Protect(password,
                    drawingObjects,
                    contents,
                    scenarios,
                    userInterfaceOnly,
                    allowFormattingCells,
                    allowFormattingColumns,
                    allowFormattingRows,
                    allowInsertingColumns,
                    allowInsertingRows,
                    allowInsertingHyperlinks,
                    allowDeletingColumns,
                    allowDeletingRows,
                    allowSorting,
                    allowFiltering,
                    allowUsingPivotTables);
            }
        }

        public virtual void Unprotect(Object password)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Unprotect(password);
            }
        }

        public void SetCellValue(int rowIndex, int columnIndex, Object value)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Cells[rowIndex, columnIndex] = value;
            }
        }

        public void Activate()
        {
            using (new EnUsCultureInvoker())
            {
                (_excelWorksheet as Microsoft.Office.Interop.Excel._Worksheet).Activate();
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Delete();
            }
        }

        public IRange GetRange(object rangeAdress)
        {
            using (new EnUsCultureInvoker())
            {
                return EntityResolver.ResolveRange(_excelWorksheet.get_Range(rangeAdress, Type.Missing));
            }
        }

        public override bool Equals(IWorksheet obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Worksheet chart = (Worksheet) obj;
            return _excelWorksheet.Equals(chart._excelWorksheet);
        }

        public void Select(bool replace)
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Select(replace);
            }
        }

        public IRange get_Range(Object cell1,Object cell2)
        {
            using (new EnUsCultureInvoker())
            {
                lock(_locker)
                {
                    return EntityResolver.ResolveRange(_excelWorksheet.get_Range(cell1, cell2));
                }
            }
        }

        public void CopyBefore(IWorksheet before)
        {
            _excelWorksheet.Copy(before.WorksheetObject, Type.Missing);
        }

        public void CopyAfter(IWorksheet after)
        {
            _excelWorksheet.Copy(Type.Missing, after.WorksheetObject);
        }

        public bool Equals(ISheet obj)
        {
            using (new EnUsCultureInvoker())
            {
				IWorksheet worksheet = obj as IWorksheet;
				if (worksheet == null)
					return false;//if obj can't be casted to IWorksheet, it definetely has different type.

				 return ((IEquatable<IWorksheet>)this).Equals(worksheet);//need to ensure correct implementation to be used.
            }
        }

        public void ClearArrows()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.ClearArrows();
            }
        }

        public Object PivotTables(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                lock (_locker)
                {
                    if (index == Type.Missing)
                    {
                        return
                            EntityResolver.ResolvePivotTables(
                                _excelWorksheet.PivotTables(Type.Missing) as Microsoft.Office.Interop.Excel.PivotTables);
                    }
                    return
                        EntityResolver.ResolvePivotTable(
                            _excelWorksheet.PivotTables(index) as Microsoft.Office.Interop.Excel.PivotTable);
                }
            }

        }

		public void SaveAs(string filename, Object fileFormat, Object password, Object writeResPassword, Object readOnlyRecommended,
							Object createBackup, Object addToMru, Object textCodepage, Object textVisualLayout, Object local)
		{
			using (new EnUsCultureInvoker())
			{
				_excelWorksheet.SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended,
							createBackup, addToMru, textCodepage, textVisualLayout, local);
			}
		}

        public void Calculate()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWorksheet.Calculate();
            }
        }
		public void Copy(IWorksheet before, IWorksheet after)
		{
			using (new EnUsCultureInvoker())
			{
				_excelWorksheet.Copy(before != null ? before.WorksheetObject : Type.Missing,
				                     after != null ? after.WorksheetObject : Type.Missing);
			}
		}
    }
}
