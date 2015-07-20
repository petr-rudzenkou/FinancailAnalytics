using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Oval : ExcelEntityWrapper<IOval>, IOval
    {
        private Microsoft.Office.Interop.Excel.Oval _excelOval;

        public Oval(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Oval oval)
            : base(entityResolver)
        {
            if (oval == null)
                throw new ArgumentNullException("oval");
            _excelOval = oval;
        }

        public IShapeRange ShapeRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapeRange(_excelOval.ShapeRange);
                }
            }
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelOval.Interior);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelOval.Border);
                }
            }
        }

        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelOval.Font);
                }
            }
        }

        public override bool Equals(IOval obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Oval oval = (Oval)obj;
            return _excelOval.Equals(oval._excelOval);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelOval);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
