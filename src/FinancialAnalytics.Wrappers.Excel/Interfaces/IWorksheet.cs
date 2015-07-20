using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    [ComVisible(true)]
    public interface IWorksheet : ISheet, IEntityWrapper<IWorksheet>
    {
        IChartObjects ChartObjects { get; }

        IApplication Application { get; }

        object WorksheetObject { get; }
        
        IShapes Shapes { get; }

        IOLEObjects OLEObjects { get; }

        INames Names { get; }

        IRange Cells { get; }

        IListObjects ListObjects { get; }
        
        ICustomProperties CustomProperties { get; }

        bool ProtectContents { get; }

        void Paste();

		void PasteSpecial(object format);

        IRange GetRange(IRange cell1, IRange cell2);

        IRange GetRange(object rangeAdress);

        void CopyBefore(IWorksheet before);

        void CopyAfter(IWorksheet after);

        void ClearArrows();

        Object PivotTables(Object index);

        IRange UsedRange { get; }

		IRange Rows { get; }

		IRange Columns { get; }

        void Protect(Object password,
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
                     bool allowUsingPivotTables);

        void Unprotect(Object password);

        IRange get_Range(Object cell1, Object cell2);

        IPageSetup PageSetup { get; }

    	void SaveAs(string filename, Object fileFormat, Object password, Object writeResPassword,
    	            Object readOnlyRecommended,
    	            Object createBackup, Object addToMru, Object textCodepage, Object textVisualLayout, Object local);

    	void Copy(IWorksheet before = null, IWorksheet after = null);

        void Paste(object value, object link);
    }
}
