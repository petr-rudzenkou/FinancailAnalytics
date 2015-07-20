using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ShapeRange : LazyEntitiesCollectionWrapper<IShapeRange, IShape>, IShapeRange
    {
        private readonly Microsoft.Office.Interop.Excel.ShapeRange _excelShapeRange;
        private readonly ExcelEntityResolver _entityResolver;

        public ShapeRange(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ShapeRange shapeRange)
        {
            if (shapeRange == null)
                throw new ArgumentNullException("shapeRange");
            _excelShapeRange = shapeRange;
            _entityResolver = entityResolver;
        }

        public override bool Equals(IShapeRange obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ShapeRange shapeRange = (ShapeRange)obj;
            return _excelShapeRange.Equals(shapeRange._excelShapeRange);
        }

        public IFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _entityResolver.ResolveFillFormat(_excelShapeRange.Fill);
                }
            }
        }


        public ILineFormat Line
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _entityResolver.ResolveLineFormat(_excelShapeRange.Line);
                }
            }
        }

		public AutoShapeType AutoShapeType
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return MsoAutoShapeTypeToAutoShapeTypeConverter.Convert(_excelShapeRange.AutoShapeType);
				}
			}
		}

        protected override void InitializeCollection()
        {
            base.InitializeCollection();
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelShapeRange.Count; i++)
                {
                    IShape shape = _entityResolver.ResolveShape(_excelShapeRange.Item(i));
                    _items.Add(shape);
                }
            }
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelShapeRange);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion		

	}
}
