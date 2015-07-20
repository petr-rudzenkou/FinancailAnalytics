using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class TickLabels : ExcelEntityWrapper<ITickLabels>, ITickLabels
    {
        protected Microsoft.Office.Interop.Excel.TickLabels _excelTickLabels;

        public TickLabels(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.TickLabels tickLabels)
            : base(entityResolver)
        {
            if (tickLabels == null)
                throw new ArgumentNullException("tickLabels");
            _excelTickLabels = tickLabels;
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
                ComObjectsFinalizer.ReleaseComObject(_excelTickLabels);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        public int Alignment
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelTickLabels.Alignment;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelTickLabels.Alignment = value;
                }
            }
        }

        public string NumberFormat
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelTickLabels.NumberFormat;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelTickLabels.NumberFormat = value;
                }
            }
        }

        public bool NumberFormatLinked
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelTickLabels.NumberFormatLinked;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelTickLabels.NumberFormatLinked = value;
                }
            }
        }

        public int Offset
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelTickLabels.Offset;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelTickLabels.Offset = value;
                }
            }
        }

        /// <summary>
        /// Returns a Font object that represents the font of the specified object.
        /// </summary>
        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelTickLabels.Font);
                }
            }
        }

        public override bool Equals(ITickLabels obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            TickLabels tickLabels = (TickLabels)obj;
            return _excelTickLabels.Equals(tickLabels);
        }



		public IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelTickLabels.Format);
				}
			}
		}
	}
}
