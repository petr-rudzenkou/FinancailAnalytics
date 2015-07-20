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
    class DropLines : ExcelEntityWrapper<IDropLines>, IDropLines
    {
        protected Microsoft.Office.Interop.Excel.DropLines _excelDropLines;

        public DropLines(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DropLines excelDropLines)
			: base(entityResolver)
		{
            if (excelDropLines == null)
			{
                throw new ArgumentNullException("excelDropLines");
			}
            _excelDropLines = excelDropLines;
		}

		public string Name
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
                    return _excelDropLines.Name;
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
                    return EntityResolver.ResolveBorder(_excelDropLines.Border);
				}
			}
		}

        public override bool Equals(IDropLines obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
            DropLines dropLines = (DropLines)obj;
            return _excelDropLines.Equals(dropLines._excelDropLines);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelDropLines);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion
    }
}
