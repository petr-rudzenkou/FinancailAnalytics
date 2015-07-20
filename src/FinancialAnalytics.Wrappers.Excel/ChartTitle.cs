using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ChartTitle : ExcelEntityWrapper<IChartTitle>, IChartTitle
    {
        protected Microsoft.Office.Interop.Excel.ChartTitle _excelChartTitle;

        public ChartTitle(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartTitle chartTitle)
            : base(entityResolver)
        {
            if (chartTitle == null)
                throw new ArgumentNullException("chartTitle");
            _excelChartTitle = chartTitle;
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
				ComObjectsFinalizer.ReleaseComObject(_excelChartTitle);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public string Text
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartTitle.Text;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartTitle.Text = value;
                }
            }
        }

        public object Orientation
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartTitle.Orientation;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartTitle.Orientation = value;
                }
            }            
        }

        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartTitle.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartTitle.Left = value;
                }
            }                
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartTitle.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartTitle.Top = value;
                }
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartTitle.Delete();
            }
        }

        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelChartTitle.Font);
                }
            }
        }

        public override bool Equals(IChartTitle obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartTitle chartTitle = (ChartTitle)obj;
            return _excelChartTitle.Equals(chartTitle._excelChartTitle);
        }
		
		public IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelChartTitle.Format);
				}
			}
		}


		public IChartFillFormat Fill
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFillFormat(_excelChartTitle.Fill); 
				}
			}
		}


        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelChartTitle.Border);
                }
            }
        }
    }
}
