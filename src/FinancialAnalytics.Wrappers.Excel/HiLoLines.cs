using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class HiLoLines : ExcelEntityWrapper<IHiLoLines>, IHiLoLines
	{
		protected Microsoft.Office.Interop.Excel.HiLoLines _excelHiLoLines;

		public HiLoLines(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.HiLoLines excelHiLoLines)
			: base(entityResolver)
		{
			if (excelHiLoLines == null)
			{
				throw new ArgumentNullException("excelHiLoLines");
			}
			_excelHiLoLines = excelHiLoLines;
		}

		public string Name
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelHiLoLines.Name;
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelHiLoLines.Border);
				}
			}
		}

		public override bool Equals(IHiLoLines obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			HiLoLines hiLoLines = (HiLoLines)obj;
			return _excelHiLoLines.Equals(hiLoLines._excelHiLoLines);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelHiLoLines);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
