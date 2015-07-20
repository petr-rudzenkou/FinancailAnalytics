using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Rectangle : ExcelEntityWrapper<IRectangle>, IRectangle
    {
        private Microsoft.Office.Interop.Excel.Rectangle _excelRectangle;

        public Rectangle(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Rectangle rectangle)
            : base(entityResolver)
        {
            if (rectangle == null)
                throw new ArgumentNullException("rectangle");
            _excelRectangle = rectangle;
        }

        public IShapeRange ShapeRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapeRange(_excelRectangle.ShapeRange);
                }
            }
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelRectangle.Interior);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelRectangle.Border);
                }
            }
        }

        public IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelRectangle.Font);
                }
            }
        }

		public double Height
		{
			get 
			{
				using (new EnUsCultureInvoker())
				{
					return _excelRectangle.Height;
				}
			}
		}

		public double Width
		{
			get 
			{
				using (new EnUsCultureInvoker())
				{
					return _excelRectangle.Width;
				}
			}
		}

        public override bool Equals(IRectangle obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Rectangle rectangle = (Rectangle)obj;
            return _excelRectangle.Equals(rectangle._excelRectangle);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelRectangle);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
