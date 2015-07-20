using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using IPivotCell = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotCell;
using IRange = FinancialAnalytics.Wrappers.Excel.Interfaces.IRange;
using IPivotField = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotField;
using IPivotItem = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotItem;
using IPivotItemList = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotItemList;
using IPivotTable = FinancialAnalytics.Wrappers.Excel.Interfaces.IPivotTable;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotCell : ExcelEntityWrapper<IPivotCell>, IPivotCell
    {
        private readonly Microsoft.Office.Interop.Excel.PivotCell _excelPivotCell;
        private static readonly XlPivotCellTypeToPivotCellTypeConverter _xlPivotCellTypeToPivotCellTypeConverter = new XlPivotCellTypeToPivotCellTypeConverter();

        public PivotCell(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotCell pivotCell)
            : base(entityResolver)
        {
            if (pivotCell == null)
            {
                throw new ArgumentNullException("pivotCell");
            }
            _excelPivotCell = pivotCell;
        }

        #region Implementation of IPivotCell

        public IApplication Application
        {
            get
            {
                using (new EnUsCultureInvoker())
                { return EntityResolver.ResolveApplication(); }
            }
        }

        public PivotCellType PivotCellType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _xlPivotCellTypeToPivotCellTypeConverter.Convert(_excelPivotCell.PivotCellType);
                }
            }
        }

        public IPivotField PivotField
        {
            get { using (new EnUsCultureInvoker())
            {
                return this.EntityResolver.ResolvePivotField(_excelPivotCell.PivotField);
            }}
        }

        public IPivotItem PivotItem
        {
            get { using (new EnUsCultureInvoker())
            {
                return this.EntityResolver.ResolvePivotItem(_excelPivotCell.PivotItem);
            }}
        }

        public IPivotTable PivotTable
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return this.EntityResolver.ResolvePivotTable(_excelPivotCell.PivotTable);
                }
            }
        }

        public IRange Range
        {
            get { using (new EnUsCultureInvoker())
            {
                return this.EntityResolver.ResolveRange(_excelPivotCell.Range);
            }}
        }

        public IPivotItemList RowItems 
        { 
            get
            {
                 using (new EnUsCultureInvoker())
                 {
                     return EntityResolver.ResolvePivotItemList(_excelPivotCell.RowItems);
                 }
            }
        }

        public IPivotItemList ColumnItems
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolvePivotItemList(_excelPivotCell.ColumnItems);
                }
            }
        }

        public IPivotField DataField
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolvePivotField(_excelPivotCell.DataField);
                }
            }
        }

        #endregion

        #region Overrides of EntityWrapperBase<IPivotCell>

        public override bool Equals(IPivotCell obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotCell pivotCell = (PivotCell)obj;
            return _excelPivotCell.Equals(pivotCell._excelPivotCell);
        }

        #endregion

        #region Disposable pattern

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelPivotCell);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
