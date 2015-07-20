using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class AxisTitle : ExcelEntityWrapper<IAxisTitle>, IAxisTitle
    {
        private Microsoft.Office.Interop.Excel.AxisTitle _excelAxisTitle;

        public AxisTitle(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.AxisTitle axisTitle)
            : base(entityResolver)
        {
            if (axisTitle == null)
                throw new ArgumentNullException("axisTitle");
            _excelAxisTitle = axisTitle;  
        }

        public string Text
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxisTitle.Text;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxisTitle.Text = value;
                }
            }
        }

        public double Top
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxisTitle.Top;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxisTitle.Top = value;
                }
            }
        }


        public double Left
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxisTitle.Left;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxisTitle.Left = value;
                }
            }
        }

        public object Orientation
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxisTitle.Orientation;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxisTitle.Orientation = value;
                }                
            }
        }

        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelAxisTitle.Font);
                }
            }
        }

        public override bool Equals(IAxisTitle obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            AxisTitle axisTitle = (AxisTitle)obj;
            return _excelAxisTitle.Equals(axisTitle._excelAxisTitle);
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
                ComObjectsFinalizer.ReleaseComObject(_excelAxisTitle);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion


		public IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelAxisTitle.Format);
				}
			}
		}

        public IBorder Border
        {
            get 
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelAxisTitle.Border);
                }
            }
        }
    }
}
