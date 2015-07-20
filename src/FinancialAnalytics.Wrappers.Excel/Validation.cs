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
using IValidation = FinancialAnalytics.Wrappers.Excel.Interfaces.IValidation;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class Validation : ExcelEntityWrapper<IValidation>, IValidation
    {
        private readonly Microsoft.Office.Interop.Excel.Validation _excelValidation;
        //private static readonly XlPivotCellTypeToPivotCellTypeConverter _xlPivotCellTypeToPivotCellTypeConverter = new XlPivotCellTypeToPivotCellTypeConverter();

        public Validation(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Validation validation)
            : base(entityResolver)
        {
            if (validation == null)
            {
                throw new ArgumentNullException("validation");
            }
            _excelValidation = validation;
        }

        #region Overrides of EntityWrapperBase<IPivotCell>

        public override bool Equals(IValidation obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Validation validation = (Validation)obj;
            return _excelValidation.Equals(validation._excelValidation);
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
                ComObjectsFinalizer.ReleaseComObject(_excelValidation);
                _disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        #region IValidation Members

        public string ErrorMessage
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelValidation.ErrorMessage;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelValidation.ErrorMessage = value;
                }
            }
        }

        public string InputMessage
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelValidation.InputMessage;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelValidation.InputMessage = value;
                }
            }
        }

        public string InputTitle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelValidation.InputTitle;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelValidation.InputTitle = value;
                }
            }
        }

        public bool IgnoreBlank
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelValidation.IgnoreBlank;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelValidation.IgnoreBlank = value;
                }
            }
        }

        public void Add(XlDVType Type, object AlertStyle, object Operator, object Formula1, object Formula2)
        {
            using (new EnUsCultureInvoker())
            {
                _excelValidation.Add(Type, AlertStyle, Operator, Formula1, Formula2);
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelValidation.Delete();
            }
        }

        #endregion
    }
}
