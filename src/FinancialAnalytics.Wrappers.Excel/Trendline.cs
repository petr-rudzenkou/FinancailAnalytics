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
    class Trendline : ExcelEntityWrapper<ITrendline>, ITrendline
    {
        private Microsoft.Office.Interop.Excel.Trendline _excelTrendline;

        public Trendline(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Trendline trendline)
            : base(entityResolver)
        {
            if (trendline == null)
                throw new ArgumentNullException("trendline");
            _excelTrendline = trendline;
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelTrendline.Border);
                }
            }
        }

        public override bool Equals(ITrendline obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Trendline trendline = (Trendline)obj;
            return _excelTrendline.Equals(trendline._excelTrendline);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelTrendline);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
