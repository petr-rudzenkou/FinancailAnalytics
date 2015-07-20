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
    internal class Line : ExcelEntityWrapper<ILine>, ILine
    {
        private Microsoft.Office.Interop.Excel.Line _excelLine;

        public Line(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Line line)
            : base(entityResolver)
        {
            if (line == null)
                throw new ArgumentNullException("line");
            _excelLine = line;
        }

        public IShapeRange ShapeRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapeRange(_excelLine.ShapeRange);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelLine.Border);
                }
            }
        }

        public override bool Equals(ILine obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Line line = (Line)obj;
            return _excelLine.Equals(line._excelLine);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelLine);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
