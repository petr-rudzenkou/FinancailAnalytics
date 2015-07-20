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
    internal class PivotField : ExcelEntityWrapper<IPivotField>, IPivotField
    {
        private readonly Microsoft.Office.Interop.Excel.PivotField _excelPivotField;

        private readonly LateBindingInvoker _invoker;
        private static Object _locker = new object();

        public PivotField(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PivotField pivotField) : base(entityResolver)
        {
            if (pivotField == null)
                throw new ArgumentNullException("pivotField");
            _excelPivotField = pivotField;
            _invoker = new LateBindingInvoker(_excelPivotField);
        }

        public PivotFieldOrientation Orientation 
        { 
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlPivotFieldOrientationToPivotFieldOrientationConverter.Convert(_excelPivotField.Orientation);
                }
            } 
            set
            {
                using (new EnUsCultureInvoker())    
                {
                    _excelPivotField.Orientation =
                   XlPivotFieldOrientationToPivotFieldOrientationConverter.ConvertBack(value);
                }
            } 
        }

        public IRange DataRange 
        { 
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelPivotField.DataRange);

                }
            }
        }

        public Object VisibleItemsList
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _invoker.InvokeGetPropertyValue<Object>("VisibleItemsList");
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _invoker.InvokeSetPropertyValue("VisibleItemsList", value);
                }
            }
        }

        public string Value
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotField.Value;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotField.Value = value;
                }
            }
        }

        public ICubeField CubeField
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveCubeField(_excelPivotField.CubeField);
                }
            }
        }

        public string CurrentPageName
        {
            get
            {
                lock(_locker)
                {
                    using (new EnUsCultureInvoker())
                    {
                        return _excelPivotField.CurrentPageName;
                    }
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotField.CurrentPageName = value;
                }
            }
        }

        public Object CurrentPageList
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotField.CurrentPageList;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotField.CurrentPageList = value;
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotField.Name;
                }
            }
        }

        public object Position
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPivotField.Position;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPivotField.Position = value;
                }
            }
        }


        public override bool Equals(IPivotField obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PivotField pivotField = (PivotField)obj;
            return _excelPivotField.Equals(pivotField._excelPivotField);
        }

        public Object get_VisibleItems(Object index)
        {
            using (new EnUsCultureInvoker())
            {
                Object visibleItemList = _excelPivotField.get_VisibleItems(index);
                if (visibleItemList is Microsoft.Office.Interop.Excel.PivotItems)
                {
                    return EntityResolver.ResolvePivotItems(visibleItemList as Microsoft.Office.Interop.Excel.PivotItems);
                }
                return EntityResolver.ResolvePivotItem(visibleItemList as Microsoft.Office.Interop.Excel.PivotItem);
            }
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
                ComObjectsFinalizer.ReleaseComObject(_excelPivotField);
                _disposed = true;
            }
            base.Dispose(disposing);
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
    }
}
