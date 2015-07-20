using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PlotArea : ExcelEntityWrapper<IPlotArea>, IPlotArea
    {
        protected Microsoft.Office.Interop.Excel.PlotArea _excelPlotArea;

        public PlotArea(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PlotArea plotArea)
            : base(entityResolver)
        {
            if (plotArea == null)
                throw new ArgumentNullException("plotArea");
            _excelPlotArea = plotArea;
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
				ComObjectsFinalizer.ReleaseComObject(_excelPlotArea);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public double Width
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPlotArea.Width;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPlotArea.Width = value;
                }
            }
        }

        public double Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPlotArea.Height;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPlotArea.Height = value;
                }
            }
        }

        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPlotArea.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPlotArea.Left = value;
                }
            } 
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPlotArea.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPlotArea.Top = value;
                }
            }
        }

        public IChart Chart
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChart(_excelPlotArea.Parent as Microsoft.Office.Interop.Excel.Chart);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelPlotArea.Border);
                }
            } 
        }

        
        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelPlotArea.Interior);
                }
            }
        }

        public IChartFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFillFormat(_excelPlotArea.Fill);
                }
            }
        }

        public object Select()
        {
            using (new EnUsCultureInvoker())
            {
                return _excelPlotArea.Select();
            }
        }

        public override bool Equals(IPlotArea obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            PlotArea plotArea = (PlotArea)obj;
            return _excelPlotArea.Equals(plotArea);
        }

		public IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelPlotArea.Format);
				}
			}
		}
	}
}
