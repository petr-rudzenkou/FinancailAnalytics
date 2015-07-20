using System;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IRange : IEntityWrapper<IRange>, System.Collections.Generic.IEnumerable<IRange>
    {
		/// <summary>
		/// Activates the range
		/// </summary>
		void Activate();

        /// <summary>
        /// Checks the range for empty status
        /// </summary>
        /// <returns>Empty status</returns>
        bool IsEmpty { get; }

        /// <summary>
        /// Get maximum rows count for worksheet
        /// </summary>
        /// <returns>Maximum rows count</returns>
        int MaxRowsCount { get; }

        /// <summary>
        /// Get maximum columns count for worksheet
        /// </summary>
        /// <returns>Maximum columns count</returns>
        int MaxColumnsCount { get; }

        /// <summary>
        /// Returns or sets the width of all columns in the specified range.
        /// If all columns in the range have the same width, the ColumnWidth property returns the width.
        /// If columns in the range have different widths, this property returns null.
        /// </summary>
        float ColumnWidth { get; set; }

        double ColumnWidthOriginal { get; set; }

        /// <summary>
        /// The width of the range in pixels.
        /// </summary>
        int Width { get; }

        /// <summary>
        /// The height of the range in pixels.
        /// </summary>
        int Height { get; }

        /// <summary>
        /// This property returns empty string if all cells in the specified range don't have the same number format.
        /// The format code is the same string as the Format Codes option in the Format Cells dialog box.
        /// The Format function uses different format code strings than do the NumberFormat and NumberFormatLocal properties.
        /// </summary>
        string NumberFormat { get; set; }

        /// <summary>
        /// Returns or sets the format code for the object as a string in the language of the user.
        /// The Format function uses different format code strings than do the NumberFormat and NumberFormatLocal properties.
        /// </summary>
        string NumberFormatLocal { get; set; }

        /// <summary>
        /// Returns or sets the vertical alignment of the specified object.
        /// </summary>
        VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// Returns or sets the horizontal alignment of the specified object.
        /// </summary>
        HorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>
        /// Determines if the rows or columns are hidden.
        /// This property returns "true" if the rows or columns are hidden.
        /// The specified range must span an entire column or row.
        /// </summary>
        bool Hidden { get; set; }

        /// <summary>
        /// This property returns "true" if Microsoft Excel wraps the text in the object and "null"
        /// if the specified range contains some cells that wrap text and other cells that don’t.
        /// Microsoft Excel will change the row height of the range, if necessary, to accommodate the text in the range.
        /// </summary>
        bool WrapText { get; set; }

        /// <summary>
        /// Returns a Font object that represents the font of the specified object.
        /// </summary>
        IFont Font { get; }

        IRange this[int rowIndex, int columnIndex] { get; }
        IRange this[object rowIndex, object columnIndex] { get;}

        void SetValue(string value);

        void SetValue(IRange value);

        object RangeObject { get; }

        IRange Cells { get; }

        IRange Rows { get; }

        IRange Columns { get; }

        int Count { get; }

        int NotEmptyCellsCount { get; }

        int Row { get; }

        int Column { get; }

        string Text { get; }

        object Delete();

        object Delete(Object shift);

        Array Values { get; }

        IWorksheet Worksheet { get; }


        IListObject ListObject { get; }

        void Select();

        Array Transpose(Array values);

        void SetValue(object value);

        IName Name { get; }

        IApplication Application { get; }

        string Address { get; }

        void Copy();

        void Copy(IRange destination);

        void CopyPicture(PictureAppearance appearance, CopyPictureFormat format);

        void PasteSpecial(PasteType pasteType, PasteSpecialOperation operation, bool skipBlanks, bool transpose);

        IRange Offset(int rowOffset, int columnOffset);

        IRange GetRange(int startRow, int startColumn, int endRow, int endColumn);

        bool Solid { get; }
		
		string ID { get; }
        
		bool IsCountValid { get; }

        string AltAddress { get; }

        /// <summary>
        /// Creates a merged cell from the specified Range object.
        /// </summary>
        void Merge();

        object AutoFit();

        IBorders Borders { get; }

        object IndentLevel { get; set; }

        void InsertIndent(int insertAmount);

        IInterior Interior { get; }

        bool HasMergeCells { get; }

        object MergeCells { get; set; }

		IRange MergeArea { get; }
		
        string Formula { get; set; }

        object FormulaObject { get; set; }

        object FormulaArray { get; set; }

		object FormulaLocal { get; set; }

        IPivotTable PivotTable { get; }

        void UnMerge();

        float RowHeight { get; set; }

		double RowHeightOriginal { get; set; }

		void ShowNavigateArrow(bool towardPrecedent);

        bool? HasFormula { get; }

        void HideNavigateArrow(bool towardPrecedent);

        IRange NavigateArrow(bool towardPrecedent, int arrowNumber, int linkNumber);

        IRange Precedents { get; }

        IRange CurrentArray { get; }

        object HasArray { get; }

        IRange Resize(object rowSize, object columnSize);

        IPivotCell PivotCell { get; }

        IRange EntireColumn { get; }

        IRange EntireRow { get; }

        Object Value2 { get; set; }

        void Copy(Object destination);

        void Insert(object Shift = null, object CopyOrigin = null);

        void Clear();

		int ReadingOrder { get; set; }

        object Orientation { get; set; }

        bool ShrinkToFit { get; set; }

		bool AddIndent { get; set; }

        bool Replace(object what, object replacement);

        bool Replace(object what, object replacement, object LookAt = null, object SearchOrder = null, object MatchCase = null, object MatchByte = null, object SearchFormat = null, object ReplaceFormat = null);

        String get_AddressLocal(Object rowAbsolute, Object columnAbsolute, ReferenceStyle refStyle, Object external,
                                Object relativeTo);

		/// <summary>
		/// Returns a Range object that represents all the cells that match the specified type
		/// </summary>
		/// <param name="type">The cells to include</param>
		/// <returns>Range object</returns>
		IRange SpecialCells(CellType type);

		/// <summary>
		/// Returns a Range object that represents all the cells that match the specified type
		/// </summary>
		/// <param name="type">The cells to include</param>
		/// <returns>Range object</returns>
		IRange SpecialCells(CellType type, object value);

        IRange Find(object what, object after = null, object lookIn = null, object lookAt = null, object searchOrder = null,
    	            SearchDirection searchDirection = SearchDirection.Next, object matchCase = null,
    	            object matchByte = null, object searchFormat = null);

    	IRange FindNext(IRange after = null);

		double Left { get; }
		double Top { get; }
        double OutlineLevel { get; }

        /// <summary>
        /// Returns validation for the range
        /// </summary>
        IValidation Validation { get; }

		bool Locked { get; set;}

		bool FormulaHidden { get; set; }

		void ExportAsFixedFormat(FixedFormatType formatType, string fileName);

        IRange get_Resize(object rowSize, object columnSize);
    }
}
