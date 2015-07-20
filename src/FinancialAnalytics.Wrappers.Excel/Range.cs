using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using System.Collections.Generic;

namespace FinancialAnalytics.Wrappers.Excel
{
	[DebuggerDisplay("{DebuggerDisplay,nq}")]
	public class Range : ExcelEntityWrapper<IRange>, IRange
	{
        private bool _disposed;
		private Microsoft.Office.Interop.Excel.Range _underlyingObject;
	    //private readonly LateBindingInvoker _invoker;

		public Range(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Range range)
			: base(entityResolver)
		{
			if (range == null)
				throw new ArgumentNullException("range");
			_underlyingObject = range;
            //_invoker = new LateBindingInvoker(_excelRange);
		}

		private string DebuggerDisplay
		{
			get
			{
				try
				{
					return string.Format("Range: {0}", Address);
				}
				catch (Exception)
				{
					return "Range: ???";
				}
			}
		}

		public object IndentLevel
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.IndentLevel;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.IndentLevel = value;
				}
			}
		}

		public void InsertIndent(int insertAmount)
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.InsertIndent(insertAmount);
			}
		}

		public IRange this[int rowIndex, int columnIndex]
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return
						EntityResolver.ResolveRange(
							_underlyingObject[rowIndex, columnIndex] as Microsoft.Office.Interop.Excel.Range);
				}
			}
			//set { _excelRange[rowIndex, columnIndex] = value.RangeObject; }
		}

		// This overload of index property need to refer to entire row or entire column.
		// In this case we can to put Type.Missing instead unused row or column index
		public IRange this[object rowIndex, object columnIndex]
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return
						// To refer to entire row or entire column to put first column index and then row index
						EntityResolver.ResolveRange(
                            _underlyingObject[rowIndex, columnIndex] as Microsoft.Office.Interop.Excel.Range);
				}
			}
			//set { _excelRange[rowIndex, columnIndex] = value.RangeObject; }
        
		}

	    public IPivotCell PivotCell
	    {
	        get
	        {
	            using (new EnUsCultureInvoker())
	            {
                    return EntityResolver.ResolvePivotCell(_underlyingObject.PivotCell);
	            }
	        }
	    }

	    public IPivotTable PivotTable
	    {
	        get
	        {
	            using (new EnUsCultureInvoker())
	            {
	                return EntityResolver.ResolvePivotTable(_underlyingObject.PivotTable);
	            }
	        }
	    }

	    public Object Value2
	    {
	        get
	        {
	            using (new EnUsCultureInvoker())
	            {
	                return _underlyingObject.Value2;
	            }
	        }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.Value2 = value;
                }
            }
	    }


        public Object FormulaArray
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.FormulaArray;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.FormulaArray = value;
                }
            }
        }


		public object AutoFit()
		{
			using (new EnUsCultureInvoker())
			{
				return _underlyingObject.AutoFit();
			}
		}


		public void SetValue(string value)
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.set_Value(Type.Missing, value);
			}
		}

		public void SetValue(IRange value)
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.set_Value(Type.Missing, value.RangeObject);
			}
		}

        public object Delete(Object shift)
        {
            using (new EnUsCultureInvoker())
            {
                return _underlyingObject.Delete(shift);
            }
        }

        public object Delete()
        {
            using (new EnUsCultureInvoker())
            {
                return Delete(Type.Missing);
            }
        }

		public object RangeObject
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject;
				}
			}
		}

		public bool IsEmpty
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetIsEmpty();
				}
			}
		}

		public int MaxRowsCount
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetMaxRowsCount();
				}
			}
		}

		public int MaxColumnsCount
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetMaxColumnsCount();
				}
			}
		}

		public IRange Cells
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
                    return EntityResolver.ResolveRange(_underlyingObject.Cells);
				}
			}
		}

		public IRange Rows
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveRange(_underlyingObject.Rows);
				}
			}
		}

		public IRange Columns
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveRange(_underlyingObject.Columns);
				}
			}
		}

		public int NotEmptyCellsCount
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetNotEmptyCellsCount();
				}
			}
		}

		public int Row
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Row;
				}
			}
		}

		public int Column
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Column;
				}
			}
		}

		/// <summary>
		/// This property returns "true" if Microsoft Excel wraps the text in the object and "null"
		/// if the specified range contains some cells that wrap text and other cells that don’t.
		/// Microsoft Excel will change the row height of the range, if necessary, to accommodate the text in the range.
		/// </summary>
		public bool WrapText
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.WrapText is bool ? (bool)_underlyingObject.WrapText : false;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.WrapText = value;
				}
			}
		}

		/// <summary>
		/// Returns or sets the width of all columns in the specified range.
		/// If all columns in the range have the same width, the ColumnWidth property returns the width.
		/// If columns in the range have different widths, this property returns null.
		/// </summary>
		public float ColumnWidth
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					object columnWith = _underlyingObject.ColumnWidth;
					if (columnWith == null)
					{
						return 0;
					}
					if (columnWith is double)
					{
						return Convert.ToSingle((double)columnWith);
					}
					return (float) columnWith;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.ColumnWidth = value;
				}
			}
		}

        public double ColumnWidthOriginal
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.ColumnWidth == null ? 0 : (double)_underlyingObject.ColumnWidth;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.ColumnWidth = value;
                }
            }
        }

		/// <summary>
		/// This property returns empty string if all cells in the specified range don't have the same number format.
		/// The format code is the same string as the Format Codes option in the Format Cells dialog box.
		/// The Format function uses different format code strings than do the NumberFormat and NumberFormatLocal properties.
		/// </summary>
		public string NumberFormat
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.NumberFormat == null ? string.Empty : (string)_underlyingObject.NumberFormat;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.NumberFormat = value;
				}
			}
		}

		/// <summary>
		/// Returns or sets the format code for the object as a string in the language of the user.
		/// The Format function uses different format code strings than do the NumberFormat and NumberFormatLocal properties.
		/// </summary>
		public string NumberFormatLocal
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.NumberFormatLocal == null || _underlyingObject.NumberFormatLocal is DBNull ? string.Empty : (string)_underlyingObject.NumberFormatLocal;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.NumberFormatLocal = value;
				}
			}
		}

		/// <summary>
		/// The width of the range.
		/// </summary>
		public int Width
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Width == null ? 0 : Convert.ToInt32(_underlyingObject.Width);
				}
			}
		}

		public int Height
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Height == null ? 0 : Convert.ToInt32(_underlyingObject.Height);
				}
			}
		}

		/// <summary>
		/// Returns or sets the vertical alignment of the specified object.
		/// </summary>
		public VerticalAlignment VerticalAlignment
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlVAlignToVerticalAlignmentConverter.Convert((Microsoft.Office.Interop.Excel.XlVAlign)_underlyingObject.VerticalAlignment);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.VerticalAlignment = XlVAlignToVerticalAlignmentConverter.ConvertBack(value);
				}
			}
		}

		/// <summary>
		/// Returns or sets the horizontal alignment for the specified object.
		/// </summary>
		public HorizontalAlignment HorizontalAlignment
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					//in some cases we get Cannot convert type 'System.DBNull' to 'Microsoft.Office.Interop.Excel.XlHAlign' exception, return default alignment in this case
					if (!(_underlyingObject.HorizontalAlignment is System.DBNull))
					{
						return XlHAlignToHorizontalAlignmentConverter.Convert((Microsoft.Office.Interop.Excel.XlHAlign) _underlyingObject.HorizontalAlignment);
					}

					return HorizontalAlignment.General;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.HorizontalAlignment = XlHAlignToHorizontalAlignmentConverter.ConvertBack(value);
				}
			}
		}

		/// <summary>
		/// Determines if the rows or columns are hidden.
		/// This property returns "true" if the rows or columns are hidden.
		/// The specified range must span an entire column or row.
		/// </summary>
		public bool Hidden
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Hidden is bool ? (bool)_underlyingObject.Hidden : false;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.Hidden = value;
				}
			}
		}

		/// <summary>
		/// Creates a merged cell from the specified Range object.
		/// </summary>
		public void Merge()
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.Merge();
			}
		}

		/// <summary>
		/// Returns a Font object that represents the font of the specified object.
		/// </summary>
		public IFont Font
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveFont(_underlyingObject.Font);
				}
			}
		}

	    public IRange EntireColumn
	    {
	        get
	        {
	            using (new EnUsCultureInvoker())
	            {
	                return EntityResolver.ResolveRange(_underlyingObject.EntireColumn);
	            }
	        }
	    }

	    public IRange EntireRow
	    {
	        get
	        {
	            using (new EnUsCultureInvoker())
	            {
	                return EntityResolver.ResolveRange(_underlyingObject.EntireRow);
	            }
	        }
	    }

		public IInterior Interior
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveInterior(_underlyingObject.Interior);
				}
			}
		}

		protected virtual bool GetIsEmpty()
		{
			using (new EnUsCultureInvoker())
			{
				double notEmptyCellsCount = _underlyingObject.Application.WorksheetFunction.CountA(_underlyingObject, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing);
				return (notEmptyCellsCount == 0);
			}
		}

		protected virtual int GetMaxRowsCount()
		{
			using (new EnUsCultureInvoker())
			{
				double maxRows = _underlyingObject.Application.WorksheetFunction.CountA(_underlyingObject.EntireColumn, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing);
				maxRows += _underlyingObject.Application.WorksheetFunction.CountBlank(_underlyingObject.EntireColumn);
				maxRows /= _underlyingObject.EntireColumn.Count;
				return (int)maxRows;
			}
		}

		protected virtual int GetMaxColumnsCount()
		{
			using (new EnUsCultureInvoker())
			{
				double maxColumns = _underlyingObject.Application.WorksheetFunction.CountA(_underlyingObject.EntireRow, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					 Type.Missing, Type.Missing);
				maxColumns += _underlyingObject.Application.WorksheetFunction.CountBlank(_underlyingObject.EntireRow);
				maxColumns /= _underlyingObject.EntireRow.Count;
				return (int)maxColumns;
			}
		}

		protected virtual int GetNotEmptyCellsCount()
		{
			using (new EnUsCultureInvoker())
			{
				double notEmptyCellsCount = _underlyingObject.Application.WorksheetFunction.CountA(_underlyingObject, Type.Missing, Type.Missing, Type.Missing,
						Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
						Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
						Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
						Type.Missing, Type.Missing);
				return (int)notEmptyCellsCount;
			}
		}

		protected virtual bool IsSolid()
		{
			using (new EnUsCultureInvoker())
			{
				return
					!_underlyingObject.get_Address(Type.Missing, Type.Missing,
											 Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Type.Missing,
											 Type.Missing).ToString().Contains(',');
			}
		}

		public int Count
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					try
					{
						return _underlyingObject.Count;
					}
					catch
					{
						return MaxRowsCount * MaxColumnsCount;
					}
				}
			}
		}

		public bool Solid
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return IsSolid();
				}
			}
		}
		
		public string ID
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.ID;
				}
			}
		}

		public string Text
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Text as string;
				}
			}
		}

		public IWorksheet Worksheet
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveWorksheet(_underlyingObject.Worksheet);
				}
			}
		}

		public IListObject ListObject
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return GetListObject();
				}
			}
		}

		public void Select()
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.Select();
			}
		}

		public Array Transpose(Array values)
		{
			using (new EnUsCultureInvoker())
			{
				return (Array)_underlyingObject.Application.WorksheetFunction.Transpose(values);
			}
		}

		public void SetValue(object value)
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.set_Value(Type.Missing, value);
			}
		}

		public IName Name
		{
			// uncomment debugger attributes to hide exception
            //[DebuggerStepperBoundary]
            //[DebuggerStepThrough]
			get
			{
				using (new EnUsCultureInvoker())
				{ return EntityResolver.ResolveName(_underlyingObject.Name as Microsoft.Office.Interop.Excel.Name); }
			}
		}

		public IApplication Application
		{
			get
			{
				using (new EnUsCultureInvoker())
				{ return EntityResolver.ResolveApplication(); }
			}
		}

		public void Activate()
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.Activate();
			}
		}

		public string Address
		{
			get
			{
				using (new EnUsCultureInvoker())
				{ return _underlyingObject.get_Address(Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing); }
			}
		}

		public void CopyPicture(PictureAppearance appearance, CopyPictureFormat format)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.XlPictureAppearance xlPictureAppearance =
					XlPictureAppearanceToPictureAppearanceConverter.ConvertBack(appearance);
				Microsoft.Office.Interop.Excel.XlCopyPictureFormat xlCopyPictureFormat =
					XlCopyPictureFormatToCopyPictureFormatConverter.ConvertBack(format);
                RepeatedCopyHelper.ExecuteCopyRepeated(() => _underlyingObject.CopyPicture(xlPictureAppearance, xlCopyPictureFormat));
			}
		}

		public void Copy()
		{
			using (new EnUsCultureInvoker())
			{
				// In EX2010 Range.Copy method sometimes don't put any data to Clipboard
				RepeatedCopyHelper.ExecuteCopyRepeated(() => _underlyingObject.Copy(Type.Missing));
			}
		}

        public void Copy(Object destination)
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Copy(destination);
            }
        }

        public void Copy(IRange destination)
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Copy(destination.RangeObject);
            }
        }

        public void Insert(object Shift, object CopyOrigin)
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Insert((Shift==null)?Type.Missing:Shift, (CopyOrigin==null)?Type.Missing:CopyOrigin);
            }
        }

        public void Clear()
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Clear();
            }
        }

		public void PasteSpecial(PasteType pasteType, PasteSpecialOperation operation, bool skipBlanks, bool transpose)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.XlPasteType xlPasteType =
					XlPasteTypeToPasteTypeConverter.ConvertBack(pasteType);
				Microsoft.Office.Interop.Excel.XlPasteSpecialOperation xlOperation =
					XlPasteSpecialOperationToPasteSpecialOperationConverter.ConvertBack(operation);
				_underlyingObject.PasteSpecial(xlPasteType, xlOperation, skipBlanks, transpose);
			}
		}

		protected virtual IListObject GetListObject()
		{
			using (new EnUsCultureInvoker())
			{
				IListObject listObject = null;
				if (_underlyingObject.ListObject != null)
				{
					listObject = EntityResolver.ResolveListObject(_underlyingObject.ListObject);
				}
				return listObject;
			}
		}

		public IRange Offset(int rowOffset, int columnOffset)
		{
			using (new EnUsCultureInvoker())
			{
				return EntityResolver.ResolveRange(_underlyingObject.Offset[rowOffset, columnOffset]);
			}
		}

		public IRange GetRange(int startRow, int startColumn, int endRow, int endColumn)
		{
			using (new EnUsCultureInvoker())
			{
                // uses Range instead of get_Range
                // http://stackoverflow.com/questions/2192977/excel-get-range-missing-when-interop-assembly-is-embedded-in-net-4-0
                Microsoft.Office.Interop.Excel.Range range = _underlyingObject.Range[_underlyingObject.get_Item(startRow, startColumn),
                        _underlyingObject.get_Item(endRow, endColumn)] as Microsoft.Office.Interop.Excel.Range;
				return EntityResolver.ResolveRange(range);
			}
		}

		public Array Values
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					object value = _underlyingObject.get_Value(Type.Missing);
					Array array = value as Array;
					return array ?? new[] { value };
				}
			}
		}

		public override bool Equals(IRange obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			Range chartTitle = (Range)obj;
			return _underlyingObject.Equals(chartTitle._underlyingObject);
		}

		public bool IsCountValid
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					try
					{
#pragma warning disable 168
						int count = _underlyingObject.Count;
#pragma warning restore 168
						return true;
					}
					catch
					{
						return false;
					}
				}
			}
		}

		public string AltAddress
		{
			get
			{
				using (new EnUsCultureInvoker())
				{ return _underlyingObject.get_Address(Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value); }
			}
		}

		public IBorders Borders
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorders(_underlyingObject.Borders);
				}
			}
		}

		public bool HasMergeCells
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					if (_underlyingObject.MergeCells is bool)
					{
						return (bool)_underlyingObject.MergeCells;
					}
					return true;
				}
			}
		}

		public object MergeCells
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.MergeCells;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.MergeCells = value;
				}
			}
		}

		public IRange MergeArea
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveRange(_underlyingObject.MergeArea);
				}
			}
		}

		public string Formula
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Formula.ToString();
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.Formula = value;
				}
			}
		}

		public object FormulaObject
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Formula;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.Formula = value;
				}
			}
		}

		public object FormulaLocal
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.FormulaLocal;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.FormulaLocal = value;
				}
			}
		}

		public void UnMerge()
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.UnMerge();
			}
		}

		public float RowHeight
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return (float)_underlyingObject.RowHeight;
				}
			}

			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.RowHeight = (object)value;
				}
			}
		}

		public double RowHeightOriginal
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.RowHeight == null ? 0 : (double)_underlyingObject.RowHeight;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.RowHeight = value;
				}
			}
		}

		public void ShowNavigateArrow(bool towardPrecedent)
		{
			using (new EnUsCultureInvoker())
			{
				if (towardPrecedent)
					_underlyingObject.ShowPrecedents(false);
				else
					_underlyingObject.ShowDependents(false);
			}
		}

        public String get_AddressLocal(Object rowAbsolute, Object columnAbsolute, ReferenceStyle refStyle, Object external, Object relativeTo)
        {
            using (new EnUsCultureInvoker())
            {
                return _underlyingObject.get_AddressLocal(rowAbsolute, columnAbsolute, XlReferenceStyleToReferenceStyleConverter.ConvertBack(refStyle), external, relativeTo);
            }
        }

		public void HideNavigateArrow(bool towardPrecedent)
		{
			using (new EnUsCultureInvoker())
			{
				if (towardPrecedent)
					_underlyingObject.ShowPrecedents(true);
				else
					_underlyingObject.ShowDependents(true);
			}
		}

		public IRange NavigateArrow(bool towardPrecedent, int arrowNumber, int linkNumber)
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.Range range = _underlyingObject.NavigateArrow(towardPrecedent, arrowNumber, linkNumber) as Microsoft.Office.Interop.Excel.Range;
				return EntityResolver.ResolveRange(range);
			}
		}

		public bool? HasFormula
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.HasFormula as bool?;
				}
			}
		}

        public bool ShrinkToFit
	    {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.ShrinkToFit is bool && (bool)_underlyingObject.ShrinkToFit;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.ShrinkToFit = value;
                }
            }             	        
	    }

		public bool AddIndent
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.AddIndent is bool && (bool)_underlyingObject.AddIndent;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.AddIndent = value;
				}
			}
		}

		public int ReadingOrder
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.ReadingOrder;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.ReadingOrder = value;
				}
			}
		}

	    public object Orientation
	    {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.Orientation;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.Orientation = value;
                }
            }  	        
	    }

        public IRange CurrentArray
        {
            get { return EntityResolver.ResolveRange(_underlyingObject.CurrentArray); }
        }

        public object HasArray
        {
            get { return _underlyingObject.HasArray; }
        }

		public IRange Precedents
		{
			get { return EntityResolver.ResolveRange(_underlyingObject.Precedents); }
		}

		public IRange Resize(object rowSize, object columnSize)
		{
			return EntityResolver.ResolveRange(_underlyingObject.Resize[rowSize, columnSize]);
		}

        public IRange get_Resize(object rowSize, object columnSize)
        {
            return EntityResolver.ResolveRange(_underlyingObject.get_Resize(rowSize, columnSize));
        }

		public IRange SpecialCells(CellType type)
		{
            using (new EnUsCultureInvoker())
            {
                bool enableEvents = _underlyingObject.Application.EnableEvents;
                if (enableEvents)
                {
                    _underlyingObject.Application.EnableEvents = false;
                }
                try
                {
                    Microsoft.Office.Interop.Excel.XlCellType xlType = XlCellTypeToCellTypeConverter.Convert(type);
                    Microsoft.Office.Interop.Excel.Range range = _underlyingObject.SpecialCells(xlType, Type.Missing);
                    return EntityResolver.ResolveRange(range);
                }
                finally
                {
                    if (enableEvents)
                    {
                        _underlyingObject.Application.EnableEvents = true;
                    }
                }
            }
		}

		public IRange SpecialCells(CellType type, object value)
		{
			Microsoft.Office.Interop.Excel.XlCellType xlType = XlCellTypeToCellTypeConverter.Convert(type);
			using (new EnUsCultureInvoker())
			{
				return EntityResolver.ResolveRange(_underlyingObject.SpecialCells(xlType, value));
			}
		}

		public IRange Find(object what, object after, object lookIn, object lookAt, object searchOrder, SearchDirection searchDirection, 
			object matchCase, object matchByte, object searchFormat)
		{
			using (new EnUsCultureInvoker())
			{
				return EntityResolver.ResolveRange(_underlyingObject.Find(what, after, lookIn, lookAt, searchOrder,
				                                                    XLSearchDirectionToSearchDirectionConverter.ConvertBack(
				                                                    	searchDirection), matchCase, matchByte, searchFormat));
			}
		}

		public IRange FindNext(IRange after)
		{
			using (new EnUsCultureInvoker())
			{
				return EntityResolver.ResolveRange(_underlyingObject.FindNext(after != null ? after.RangeObject : Type.Missing));
			}
		}

		public bool Replace(object what, object replacement)
        {
            using (new EnUsCultureInvoker())
            {
                return _underlyingObject.Replace(what, replacement);
            }
        }

        public bool Replace(object what, object replacement, object LookAt, object SearchOrder, object MatchCase, object MatchByte, object SearchFormat, object ReplaceFormat)
        {
            using (new EnUsCultureInvoker())
            {
                return _underlyingObject.Replace(what, replacement, (LookAt == null) ? Type.Missing : LookAt, (SearchOrder == null) ? Type.Missing : SearchOrder, (MatchCase == null) ? Type.Missing : MatchCase, (MatchByte == null) ? Type.Missing : MatchByte, (SearchFormat == null) ? Type.Missing : SearchFormat, (ReplaceFormat == null) ? Type.Missing : ReplaceFormat);
            }
        }

        public IEnumerator<IRange> GetEnumerator()
        {
            List<IRange> cells = new List<IRange>();
            foreach (Microsoft.Office.Interop.Excel.Range cell in _underlyingObject)
            {
                cells.Add(EntityResolver.ResolveRange(cell));
            }
            return cells.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            System.Collections.ArrayList cells = new System.Collections.ArrayList();
            foreach (Microsoft.Office.Interop.Excel.Range cell in _underlyingObject)
            {
                cells.Add(EntityResolver.ResolveRange(cell));
            }
            return cells.GetEnumerator();
        }

		public double Left
		{
			get 
			{
				using (new EnUsCultureInvoker())
				{
					return (double)_underlyingObject.Left;
				}
			}
		}

		public double Top
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return (double)_underlyingObject.Top;
				}
			}
		}

        public double OutlineLevel
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return (double)_underlyingObject.OutlineLevel;
				}
			}
		}

        public IValidation Validation
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveValidation(_underlyingObject.Validation);
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_underlyingObject);
                _underlyingObject = null;
                _disposed = true;
            }
            base.Dispose(disposing);
        }


		public bool Locked
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.Locked; 
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.Locked = value; 
				}
			}
		}

		public bool FormulaHidden
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _underlyingObject.FormulaHidden;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_underlyingObject.FormulaHidden = value;
				}
			}
		}

		public void ExportAsFixedFormat(FixedFormatType formatType, string fileName)
		{
			using (new EnUsCultureInvoker())
			{
				_underlyingObject.GetType().InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, _underlyingObject, new object[] { formatType, fileName });
			}
		}
	}
}