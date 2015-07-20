using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class PageSetup : ExcelEntityWrapper<IPageSetup>, IPageSetup
    {
        #region Constants and Fields

        protected Microsoft.Office.Interop.Excel.PageSetup _excelPageSetup;

        private bool _disposed;

        #endregion

        #region Constructors and Destructors

        public PageSetup(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.PageSetup pageSetup)
            : base(entityResolver)
        {
            if (pageSetup == null)
            {
                throw new ArgumentNullException("pageSetup");
            }
            _excelPageSetup = pageSetup;
        }

        #endregion

        #region Properties

        public ObjectSize ChartSize
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlObjectSizeToObjectSizeConverter.Convert(_excelPageSetup.ChartSize);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPageSetup.ChartSize = XlObjectSizeToObjectSizeConverter.ConvertBack(value);
                }
            }
        }

        public string PrintArea
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPageSetup.PrintArea;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPageSetup.PrintArea = value;
                }
            }
		}

		public object FitToPagesTall
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.FitToPagesTall;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.FitToPagesTall = value;
				}
			}
		}

		public object FitToPagesWide
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.FitToPagesWide;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.FitToPagesWide = value;
				}
			}
		}

		public object Zoom
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.Zoom;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.Zoom = value;
				}
			}
		}

		public string LeftFooter
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.LeftFooter;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.LeftFooter = value;
				}
			}
		}

		public string CenterFooter
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.CenterFooter;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.CenterFooter = value;
				}
			}
		}

		public string RightFooter
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.RightFooter;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.RightFooter = value;
				}
			}
		}

		public string LeftHeader
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.LeftHeader;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.LeftHeader = value;
				}
			}
		}

		public string CenterHeader
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.CenterHeader;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.CenterHeader = value;
				}
			}
		}

		public string RightHeader
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.RightHeader;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.RightHeader = value;
				}
			}
		}

		public double TopMargin
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.TopMargin;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.TopMargin = value;
				}
			}
		}

		public double RightMargin
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.RightMargin;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.RightMargin = value;
				}
			}
		}

		public double BottomMargin
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.BottomMargin;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.BottomMargin = value;
				}
			}
		}

		public double LeftMargin
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPageSetup.LeftMargin;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.LeftMargin = value;
				}
			}
		}

		public PageOrientation Orientation
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlPageOrientationToPageOrientationConverter.Convert(_excelPageSetup.Orientation);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPageSetup.Orientation = XlPageOrientationToPageOrientationConverter.ConvertBack(value); ;
				}
			}
		}
		public IApplication Application
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveApplication();
				}
			}
		}

        #endregion

        #region Implemented Interfaces

        #region IPageSetup

        public override bool Equals(IPageSetup obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            var chartTitle = (PageSetup)obj;
            return _excelPageSetup.Equals(chartTitle._excelPageSetup);
        }

        #endregion

        #endregion

        #region Methods

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelPageSetup);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}