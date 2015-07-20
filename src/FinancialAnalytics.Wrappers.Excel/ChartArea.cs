using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class ChartArea : ExcelEntityWrapper<IChartArea>, IChartArea
    {
        protected Microsoft.Office.Interop.Excel.ChartArea _excelChartArea;

        public ChartArea(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartArea chartArea)
            : base(entityResolver)
        {
            if (chartArea == null)
                throw new ArgumentNullException("chartArea");
            _excelChartArea = chartArea;
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
				ComObjectsFinalizer.ReleaseComObject(_excelChartArea);
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
                    return _excelChartArea.Width;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartArea.Width = value;
                }
            }
        }

        public double Height
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartArea.Height;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartArea.Height = value;
                }
            }
        }

        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartArea.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartArea.Left = value;
                }
            } 
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartArea.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartArea.Top = value;
                }
            }
        }

        public bool Shadow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartArea.Shadow;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartArea.Shadow = value;
                }
            }
        }

        public IChart Chart
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChart(_excelChartArea.Parent as Microsoft.Office.Interop.Excel.Chart);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelChartArea.Border);
                }
            } 
        }

        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelChartArea.Font);
                }
            }
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelChartArea.Interior);
                }
            }
        }

        public IChartFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFillFormat(_excelChartArea.Fill);
                }
            }
        }

        public void Select()
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartArea.Select();
            }
        }

        public void Copy()
        {
            using (new EnUsCultureInvoker())
            {
                // In EX2010 ChartArea.Copy method sometimes don't put any data to Clipboard
                RepeatedCopyHelper.ExecuteCopyRepeated(() => _excelChartArea.Copy());
            }
        }

        public override bool Equals(IChartArea obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartArea chartArea = (ChartArea)obj;
            return _excelChartArea.Equals(chartArea._excelChartArea);
        }



		public IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelChartArea.Format);
				}
			}
		}
	}
}
