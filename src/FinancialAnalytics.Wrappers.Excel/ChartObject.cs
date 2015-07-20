using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class ChartObject : ExcelEntityWrapper<IChartObject>, IChartObject
    {
        protected Microsoft.Office.Interop.Excel.ChartObject _excelChartObject;

        public ChartObject(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartObject chartObject)
            : base(entityResolver)
        {
            if (chartObject == null)
                throw new ArgumentNullException("chartObject");
            _excelChartObject = chartObject;
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
				ComObjectsFinalizer.ReleaseComObject(_excelChartObject);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public IChart Chart
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChart(_excelChartObject.Chart);
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartObject.Name;
                }
            }
        }

        public double Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartObject.Height;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartObject.Height = value;
                }
            }
        }

        public double Width
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartObject.Width;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartObject.Width = value;
                }
            }
        }

        public void Select(bool replace)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartObject.Select(replace);
            }
        }

        public void Activate()
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartObject.Activate();
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartObject.Delete();
                Dispose();
            }
        }
        
        public override bool Equals(IChartObject obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartObject chartObject = (ChartObject)obj;
            return _excelChartObject.Equals(chartObject._excelChartObject);
        }

		public double Top
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChartObject.Top;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChartObject.Top = value;
				}
			}
		}

		public double Left
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelChartObject.Left;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelChartObject.Left = value;
				}
			}
		}

	}
}
