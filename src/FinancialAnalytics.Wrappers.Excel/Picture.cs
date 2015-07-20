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
    internal class Picture : ExcelEntityWrapper<IPicture>, IPicture
    {
        private Microsoft.Office.Interop.Excel.Picture _excelPicture;

        public Picture(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Picture picture)
            : base(entityResolver)
        {
            if (picture == null)
                throw new ArgumentNullException("picture");
            _excelPicture = picture;
        }

        public IShapeRange ShapeRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapeRange(_excelPicture.ShapeRange);
                }
            }
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelPicture.Interior);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelPicture.Border);
                }
            }
        }

        public override bool Equals(IPicture obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Picture picture = (Picture)obj;
            return _excelPicture.Equals(picture._excelPicture);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelPicture);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
