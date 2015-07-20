using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class AutoRecover : ExcelEntityWrapper<IAutoRecover>, IAutoRecover
    {
        private Microsoft.Office.Interop.Excel.AutoRecover _excelAutoRecover;

        public AutoRecover(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.AutoRecover autoRecover)
            : base(entityResolver)
        {
            if (autoRecover == null)
                throw new ArgumentNullException("autoRecover");
            _excelAutoRecover = autoRecover;
        }

        public bool Enabled
        {
            get { return _excelAutoRecover.Enabled; }
            set { _excelAutoRecover.Enabled = value; }
        }

        public int Time
        {
            get { return _excelAutoRecover.Time; }
            set { _excelAutoRecover.Time = value; }
        }

        public string Path
        {
            get { return _excelAutoRecover.Path; }
            set { _excelAutoRecover.Path = value; }
        }

        #region Disposable pattern

        private bool disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelAutoRecover);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        public override bool Equals(IAutoRecover obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            AutoRecover autoRecover = (AutoRecover)obj;
            return _excelAutoRecover.Equals(autoRecover._excelAutoRecover);
        }
    }
}
