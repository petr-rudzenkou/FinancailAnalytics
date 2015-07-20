using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotTables : EntitiesCollectionWrapperBase<IPivotTables, IPivotTable>, IPivotTables
    {
        private ExcelEntityResolver EntityResolver { get; set; }
        private readonly Microsoft.Office.Interop.Excel.PivotTables _excelPivotTables;

        public PivotTables(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotTables pivotTables)
        {
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            if (pivotTables == null)
            {
                throw new ArgumentNullException("pivotTables");
            }
            this.EntityResolver = entityResolver;
            _excelPivotTables = pivotTables;
            InitializeCollection();
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPivotTables.Count; i ++)
                {
                    IPivotTable pivotTable = EntityResolver.ResolvePivotTable(_excelPivotTables.Item(i));
                    _items.Add(pivotTable);
                }
            }
        }

        public override bool Equals(IPivotTables obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotTables pivotTables = (PivotTables)obj;
            return _excelPivotTables.Equals(pivotTables._excelPivotTables);
        }

        public IPivotTable Add(
	        IPivotCache pivotCache, 
	        Object tableDestination, 
	        Object tableName, 
	        Object readData,
            PivotTableVersionList defaultVersion)
        {
            using (new EnUsCultureInvoker())
            {
                return EntityResolver.ResolvePivotTable(_excelPivotTables.Add((Microsoft.Office.Interop.Excel.PivotCache)pivotCache.PivotCacheObject, tableDestination, tableName, readData, defaultVersion));
            }
        }

        #region Disposable pattern

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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotTables);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion
    }
}
