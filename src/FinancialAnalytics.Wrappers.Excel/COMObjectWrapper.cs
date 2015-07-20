using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class COMObjectWrapper : ExcelEntityWrapper<ICOMObject>, ICOMObject
    {
        private Object _comObject;

        public COMObjectWrapper(ExcelEntityResolver resolver, Object comObject)
            : base(resolver)
        {
            if (comObject == null)
            {
                throw new ArgumentNullException("comObject");
            }
            if (!comObject.GetType().IsCOMObject)
            {
                throw new ArgumentException("comObject is not really a COM object!");
            }
            _comObject = comObject;
        }


        public override bool Equals(ICOMObject obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            COMObjectWrapper comObject = (COMObjectWrapper)obj;
            return _comObject.Equals(comObject._comObject);
        }

        private bool _disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {

                }
                ComObjectsFinalizer.ReleaseComObject(_comObject);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
    }
}
