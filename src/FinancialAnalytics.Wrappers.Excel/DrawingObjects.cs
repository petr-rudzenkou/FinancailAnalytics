using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class DrawingObjects : ExcelEntityWrapper<IDrawingObjects>, IDrawingObjects
    {
        private Microsoft.Office.Interop.Excel.DrawingObjects _excelDrawingObjects;

        public DrawingObjects(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DrawingObjects drawingObjects)
            : base(entityResolver)
        {
            if (drawingObjects == null)
                throw new ArgumentNullException("drawingObjects");
            _excelDrawingObjects = drawingObjects;
        }

        public IShapeRange ShapeRange
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveShapeRange(_excelDrawingObjects.ShapeRange);
                }
            }
        }

        public override bool Equals(IDrawingObjects obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            DrawingObjects drawingObjects = (DrawingObjects)obj;
            return _excelDrawingObjects.Equals(drawingObjects._excelDrawingObjects);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelDrawingObjects);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
