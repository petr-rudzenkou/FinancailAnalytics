using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotFields : EntitiesCollectionWrapperBase<IPivotFields, IPivotField>, IPivotFields
    {
        protected ExcelEntityResolver EntityResolver { get; private set; }
        private readonly MSExcel.PivotFields _excelPivotFields;

        public PivotFields(ExcelEntityResolver entityResolver, MSExcel.PivotFields excelPivotFields)
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            if (excelPivotFields == null)
                throw new ArgumentNullException("excelPivotFields");
            this.EntityResolver = entityResolver;
            _excelPivotFields = excelPivotFields;
            InitializeCollection();

        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPivotFields.Count; i ++)
                {
                    AddToCollection(_excelPivotFields.Item(i));
                }
            }
        }

        private void AddToCollection(Object pivotField)
        {
            _items.Add(
                        EntityResolver.ResolvePivotField(pivotField as MSExcel.PivotField)
                        );
        }

        public override bool Equals(IPivotFields obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotFields pivotFields = (PivotFields)obj;
            return _excelPivotFields.Equals(pivotFields._excelPivotFields);
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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotFields);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
    }
}
