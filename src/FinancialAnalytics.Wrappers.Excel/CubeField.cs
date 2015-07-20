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
using MSExcel = Microsoft.Office.Interop.Excel;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class CubeField : ExcelEntityWrapper<ICubeField>, ICubeField
    {
        private readonly MSExcel.CubeField _excelCubeField;
        private readonly LateBindingInvoker _invoker;

        public CubeField(ExcelEntityResolver entityResolver, MSExcel.CubeField excelCubeField)
            : base(entityResolver)
        {
            if (excelCubeField == null)
            {
                throw new ArgumentNullException("excelCubeField");
            }
            _excelCubeField = excelCubeField;
            _invoker = new LateBindingInvoker(_excelCubeField);
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
                ComObjectsFinalizer.ReleaseComObject(_excelCubeField);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion


        public override bool Equals(ICubeField obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            CubeField cubeField = (CubeField)obj;
            return _excelCubeField.Equals(cubeField._excelCubeField);
        }

        public bool EnableMultiplePageItems
        {
            get
            {
                using(new EnUsCultureInvoker())
                {
                    return _excelCubeField.EnableMultiplePageItems;
                }
            }
            set
            {
                using(new EnUsCultureInvoker())
                {
                    _excelCubeField.EnableMultiplePageItems = value;
                }
            }
        }

        public string CurrentPageName
        {
            get
            {
                using(new EnUsCultureInvoker())
                {
                    return _invoker.InvokeGetPropertyValue<String>("CurrentPageName");
                }
            }
            set
            {
                using(new EnUsCultureInvoker())
                {
                    _invoker.InvokeSetPropertyValue("CurrentPageName", value);
                }
            }
        }

        public IPivotFields PivotFields
        {
            get
            {
                using(new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolvePivotFields(_excelCubeField.PivotFields);
                }
            }
        }

        public PivotFieldOrientation Orientation
        {
            get
            {
                using(new EnUsCultureInvoker())
                {
                    return XlPivotFieldOrientationToPivotFieldOrientationConverter.Convert(_excelCubeField.Orientation);
                }
            }
            set
            {
                using(new EnUsCultureInvoker())
                {
                    _excelCubeField.Orientation =
                        XlPivotFieldOrientationToPivotFieldOrientationConverter.ConvertBack(value);
                }
            }
        }
    }
}
