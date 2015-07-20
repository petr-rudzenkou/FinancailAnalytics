using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class ChartGroup : ExcelEntityWrapper<IChartGroup>, IChartGroup
	{
		protected Microsoft.Office.Interop.Excel.ChartGroup _excelChartGroup;

		public ChartGroup(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartGroup chartGroup)
			: base(entityResolver)
		{
			if (chartGroup == null)
			{
				throw new ArgumentNullException("chartGroup");
			}
			_excelChartGroup = chartGroup;
		}

		public bool HasHiLoLines
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChartGroup.HasHiLoLines;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChartGroup.HasHiLoLines = value;
				}
			}
		}

		public bool HasUpDownBars
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChartGroup.HasUpDownBars;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChartGroup.HasUpDownBars = value;
				}
			}
		}

		public IBars DownBars
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveDownBars(_excelChartGroup.DownBars);
				}
			}
		}

		public IBars UpBars
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveUpBars(_excelChartGroup.UpBars);
				}
			}
		}

		public IHiLoLines HiLoLines
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveHiLoLines(_excelChartGroup.HiLoLines);
				}
			}
		}

		public override bool Equals(IChartGroup obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ChartGroup chartGroup = (ChartGroup)obj;
			return _excelChartGroup.Equals(chartGroup._excelChartGroup);
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelChartGroup);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
