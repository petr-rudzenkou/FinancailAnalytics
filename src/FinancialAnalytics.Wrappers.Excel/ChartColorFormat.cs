using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ChartColorFormat : ExcelEntityWrapper<IChartColorFormat>, IChartColorFormat
    {
        protected Microsoft.Office.Interop.Excel.ChartColorFormat _excelChartColorFormat;

        public ChartColorFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartColorFormat chartColorFormat)
            : base(entityResolver)
        {
            if (chartColorFormat == null)
                throw new ArgumentNullException("chartColorFormat");
            _excelChartColorFormat = chartColorFormat;
        }

		#region Disposable pattern

		private bool disposed = false;
		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				if (disposing)
				{
					//Here we must dispose managed resources
				}
				//Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_excelChartColorFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public int RGB
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartColorFormat.RGB;
                }
            }
        }

        public int SchemeColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartColorFormat.SchemeColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartColorFormat.SchemeColor = value;
                }
            }
        }

        public override bool Equals(IChartColorFormat obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartColorFormat chartColorFormat = (ChartColorFormat)obj;
            return _excelChartColorFormat.Equals(chartColorFormat._excelChartColorFormat);
        }

    }
}
