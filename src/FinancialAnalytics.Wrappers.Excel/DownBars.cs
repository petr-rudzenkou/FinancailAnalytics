using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class DownBars : ExcelEntityWrapper<IBars>, IBars
	{
		protected Microsoft.Office.Interop.Excel.DownBars _excelBars;

		public DownBars(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DownBars excelBars)
			: base(entityResolver)
		{
			if (excelBars == null)
			{
				throw new ArgumentNullException("excelBars");
			}
			_excelBars = excelBars;
		}

		public string Name
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelBars.Name;
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelBars.Border);
				}
			}
		}

		public IInterior Interior
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveInterior(_excelBars.Interior);
				}
			}
		}

		public IChartFillFormat Fill
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFillFormat(_excelBars.Fill);
				}
			}
		}

		public override bool Equals(IBars obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			DownBars bars = (DownBars)obj;
			return _excelBars.Equals(bars._excelBars);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelBars);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
