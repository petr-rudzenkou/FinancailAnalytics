using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Name : ExcelEntityWrapper<IName>, IName
    {
        #region Constants and Fields

        protected Microsoft.Office.Interop.Excel.Name _excelName;

        private bool _disposed;

        #endregion

        #region Constructors and Destructors

        public Name(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Name name)
            : base(entityResolver)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }
            _excelName = name;
        }

        #endregion

        #region Properties

        public string NameLocal
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelName.NameLocal;
                }
            }
        }


        public string RangeName
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelName.Name;
                }
            }
        }

        public string RangeValue
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelName.Value;
                }
            }
        }

        public string RefersTo
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelName.RefersTo.ToString();
                }
            }
        }

        public IRange RefersToRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveRange(_excelName.RefersToRange);
                }
            }
        }

        public bool Visible
        {
            get
            {
                return _excelName.Visible;
            }
            set
            {
                _excelName.Visible = value;
            }
        }

        #endregion

        #region Implemented Interfaces

        #region IEquatable<IName>

        public override bool Equals(IName obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            var chartTitle = (Name)obj;
            return _excelName.Equals(chartTitle._excelName);
        }

        #endregion

        #endregion

        #region Methods

		public void Delete()
		{
			using (new EnUsCultureInvoker())
			{
				_excelName.Delete();
			}
		}

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelName);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}