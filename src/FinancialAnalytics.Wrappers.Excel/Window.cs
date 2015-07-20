using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Window : ExcelEntityWrapper<IWindow>, IWindow
    {
        private readonly Microsoft.Office.Interop.Excel.Window _excelWindow;

        public Window(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Window window)
            : base(entityResolver)
        {
            if (window == null)
                throw new ArgumentNullException("window");
            _excelWindow = window;
        }

        #region Disposable pattern

        private bool disposed;
        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelWindow);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        public double Zoom
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.Zoom;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.Zoom = value;
                }
            }
        }

        public bool DisplayGridlines
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.DisplayGridlines;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.DisplayGridlines = value;
                }
            }
        }

        public int ScrollRow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.ScrollRow;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.ScrollRow = value;
                }
            }
        }

        public int ScrollColumn
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.ScrollColumn;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.ScrollColumn = value;
                }
            }
        }

        public int SplitRow
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.SplitRow;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.SplitRow = value;
                }
            }
        }

        public int SplitColumn
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.SplitColumn;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.SplitColumn = value;
                }
            }
        }

        public bool FreezePanes
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.FreezePanes;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.FreezePanes = value;
                }
            }
        }


        public override bool Equals(IWindow obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Window chartTitle = (Window)obj;
            return _excelWindow.Equals(chartTitle._excelWindow);
        }

        public void Activate()
        {
            using (new EnUsCultureInvoker())
            {
                _excelWindow.Activate();
            }
        }

        public bool Visible
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.Visible;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelWindow.Visible = value;
                }
            }
        }

        public object WindowObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow;
                }
            }
        }

        public string Caption
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelWindow.Caption != null ?_excelWindow.Caption.ToString() : string.Empty;
                }
            }
        }

		public ISheets Sheets
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveSheets(_excelWindow.SelectedSheets);
				}
			}
		}

		public IRange VisibleRange
		{
			get 
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveRange(_excelWindow.VisibleRange);
				}
			}
		}

        public IPanes Panes
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolvePanes(_excelWindow.Panes);
                }
            }
        }

        public int PointsToScreenPixelsX(int points)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelWindow.PointsToScreenPixelsX(points);
            }
        }

        public int PointsToScreenPixelsY(int points)
        {
            using (new EnUsCultureInvoker())
            {
                return _excelWindow.PointsToScreenPixelsY(points);
            }
        }

		public void ScrollIntoView(int left, int top, int width, int height)
		{
			using (new EnUsCultureInvoker())
			{
				_excelWindow.ScrollIntoView(left, top, width, height, true);
			}
		}

		public ISheet ActiveSheet
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveSheet(_excelWindow.ActiveSheet);
				}
			}
		}
	}
}
