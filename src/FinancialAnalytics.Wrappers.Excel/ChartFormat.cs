using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class ChartFormat : ExcelEntityWrapper<Interfaces.IChartFormat>, Interfaces.IChartFormat
	{
		#region Constants and Fields

		protected Microsoft.Office.Interop.Excel.ChartFormat _chartFormat;
		private bool _disposed;

		#endregion

		public ChartFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartFormat chartFormat)
			: base(entityResolver)
		{
			if (chartFormat == null)
			{
				throw new ArgumentNullException("chartFormat");
			}
			_chartFormat = chartFormat;
			
		}

		public override bool Equals(Interfaces.IChartFormat obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			var chartFormat = (ChartFormat)obj;
			return _chartFormat.Equals(chartFormat._chartFormat);
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
				ComObjectsFinalizer.ReleaseComObject(_chartFormat);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		public Interfaces.ITextFrame2 TextFrame2
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveTextFrame2(_chartFormat.TextFrame2);
				}
			}
		}

		public Interfaces.IFillFormat Fill
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveFillFormat(_chartFormat.Fill);
				}
			}
		}

		public Interfaces.ILineFormat Line
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveLineFormat(_chartFormat.Line); 
				}
			}
		}
	}
}