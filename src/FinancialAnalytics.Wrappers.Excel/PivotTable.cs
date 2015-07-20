using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PivotTable : ExcelEntityWrapper<IPivotTable>, IPivotTable
    {
        private readonly Microsoft.Office.Interop.Excel.PivotTable _excelPivotTable;
        private readonly LateBindingInvoker _invoker;
        public PivotTable(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotTable pivotTable)
            : base(entityResolver)
        {
            if (pivotTable == null)
                throw new ArgumentNullException("pivotTable");
            _excelPivotTable = pivotTable;
            _invoker = new LateBindingInvoker(_excelPivotTable);
        }

        public override bool Equals(IPivotTable obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotTable pivotTable = (PivotTable)obj;
            return _excelPivotTable.Equals(pivotTable._excelPivotTable);
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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotTable);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #region IPivotTable Members

        public IPivotCache PivotCache()
        {
             using (new EnUsCultureInvoker())
             {
                 Microsoft.Office.Interop.Excel.PivotCache excelPivotCache = _excelPivotTable.PivotCache();
                 IPivotCache result = this.EntityResolver.ResolvePivotCache(excelPivotCache);
                 return result;
             }
        }

        public Object get_DataFields(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object retValue = _excelPivotTable.get_DataFields(index);
                if (retValue is Microsoft.Office.Interop.Excel.PivotFields)
                {
                    return EntityResolver.ResolvePivotFields(retValue as Microsoft.Office.Interop.Excel.PivotFields);
                }
                return EntityResolver.ResolvePivotField(retValue as Microsoft.Office.Interop.Excel.PivotField);
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotTable.Name;
                }
            }
        }

        public PivotTableVersionList Version
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlPivotTableVersionListToPivotTableVersionListConverter.Convert(_excelPivotTable.Version);
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

        public ICubeFields CubeFields
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveCubeFields(_excelPivotTable.CubeFields);
                }
            }
        }

        public bool EnableDrilldown
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotTable.EnableDrilldown;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotTable.EnableDrilldown = value;
                }
            }
        }

        public IRange TableRange2
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelPivotTable.TableRange2);
                }
            }
        }

        public bool ShowTableStyleRowStripes
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _invoker.InvokeGetPropertyValue<bool>("ShowTableStyleRowStripes");
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _invoker.InvokeSetPropertyValue("ShowTableStyleRowStripes", value);
                }
            }
        }

        public bool ManualUpdate
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotTable.ManualUpdate;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotTable.ManualUpdate = value;
                }
            }
        }

        public Object TableStyle2
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _invoker.InvokeGetPropertyValue<Object>("TableStyle2");
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _invoker.InvokeSetPropertyValue("TableStyle2", value);
                }
            }
        }

        public Object Parent
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotTable.Parent;
                }
            }
        }

        public IWorksheet ParentWorksheet 
        { 
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return (_excelPivotTable.Parent is Microsoft.Office.Interop.Excel.Worksheet
                                ? EntityResolver.ResolveWorksheet(_excelPivotTable.Parent as Microsoft.Office.Interop.Excel.Worksheet)
                                : null);
                }
            }
        }

        public IRange RowRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelPivotTable.RowRange);
                }
            }
        }

        public void Update()
        {
             using (new EnUsCultureInvoker())
             {
                 _excelPivotTable.Update();
             }
        }

        public Object PivotFields(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object excelPivotFields = _excelPivotTable.PivotFields(index);
                if (excelPivotFields is Microsoft.Office.Interop.Excel.PivotFields)
                {
                    return EntityResolver.ResolvePivotFields(excelPivotFields as Microsoft.Office.Interop.Excel.PivotFields);
                }
                return EntityResolver.ResolvePivotField(excelPivotFields as Microsoft.Office.Interop.Excel.PivotField);
            }
        }

        public Object get_ColumnFields(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object columnFields = _excelPivotTable.get_ColumnFields(index);
                if (columnFields is Microsoft.Office.Interop.Excel.PivotFields)
                {
                    return EntityResolver.ResolvePivotFields(
                        columnFields as Microsoft.Office.Interop.Excel.PivotFields
                        );
                }
                return EntityResolver.ResolvePivotField(columnFields as Microsoft.Office.Interop.Excel.PivotField);
            }
        }

        public Object get_RowFields(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object rowFields = _excelPivotTable.get_RowFields(index);
                if (rowFields is Microsoft.Office.Interop.Excel.PivotFields)
                {
                    return EntityResolver.ResolvePivotFields(
                        rowFields as Microsoft.Office.Interop.Excel.PivotFields
                        );
                }
                return EntityResolver.ResolvePivotField(rowFields as Microsoft.Office.Interop.Excel.PivotField);
            }
        }

        public Object get_PageFields(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object pageFields = _excelPivotTable.get_PageFields(index);
                if (pageFields is Microsoft.Office.Interop.Excel.PivotFields)
                {
                    return EntityResolver.ResolvePivotFields(pageFields as Microsoft.Office.Interop.Excel.PivotFields);
                }
                return EntityResolver.ResolvePivotField(pageFields as Microsoft.Office.Interop.Excel.PivotField);
            }
        }

        public IRange TableRange1
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelPivotTable.TableRange1);
                }
            }
        }

        #endregion
    }
}
