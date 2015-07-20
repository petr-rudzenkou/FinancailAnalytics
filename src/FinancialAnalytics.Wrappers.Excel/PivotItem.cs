using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotItem : ExcelEntityWrapper<IPivotItem>, IPivotItem
    {
        private readonly Microsoft.Office.Interop.Excel.PivotItem _excelPivotItem;

        public PivotItem(ExcelEntityResolver excelEntityResolver, Microsoft.Office.Interop.Excel.PivotItem pivotItem) 
            : base(excelEntityResolver)
        {
            if (pivotItem == null)
            {
                throw new ArgumentNullException("pivotItem");
            }
            _excelPivotItem = pivotItem;
        }

        

        public override bool Equals(IPivotItem obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotItem pivotItem = (PivotItem)obj;
            return _excelPivotItem.Equals(pivotItem._excelPivotItem);
        }

        private bool _disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelPivotItem);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        public string Value
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotItem.Value;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotItem.Value = value;
                }
            }
        }

        public int Position
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotItem.Position;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotItem.Position = value;
                }
            }

        }

        public bool DrilledDown
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotItem.DrilledDown;
                }
            }
            set 
            { 
                using (new EnUsCultureInvoker())
                {
                    _excelPivotItem.DrilledDown = value;
                } 
            }
        }
    }
}
