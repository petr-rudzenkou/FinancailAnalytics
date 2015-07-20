using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class Shapes : LazyEntitiesCollectionWrapper<IShapes, IShape>, IShapes
    {
        protected ExcelEntityResolver _entityResolver;
        Microsoft.Office.Interop.Excel.Shapes _excelShapes;

        public Shapes(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Shapes shapes)
        {
            if (shapes == null)
                throw new ArgumentNullException("shapes");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelShapes = shapes;
            _entityResolver = entityResolver;
        }

		public override int Count
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelShapes.Count;
				}
			}
		}

		public IShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument, float left, float top, float width = -1f, float height = -1f)
		{
			using (new EnUsCultureInvoker())
			{
				return _entityResolver.ResolveShape(_excelShapes.AddPicture(fileName, MsoTriStateToBoolConverter.ConvertBack(linkToFile),
					                                                     MsoTriStateToBoolConverter.ConvertBack(saveWithDocument), left,
					                                                     top, width, height));
			}
		}

		public IShape AddShape(AutoShapeType shapeType, float left, float top, float width, float height)
		{
			using (new EnUsCultureInvoker())
			{
				return _entityResolver.ResolveShape(_excelShapes.AddShape(MsoAutoShapeTypeToAutoShapeTypeConverter.ConvertBack(shapeType), left, top, width, height));
			}
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
				ComObjectsFinalizer.ReleaseComObject(_excelShapes);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

		protected override void InitializeCollection()
		{
			base.InitializeCollection();
			using (new EnUsCultureInvoker())
			{
				for (int i = 1; i <= _excelShapes.Count; i++)
				{
					IShape shape = _entityResolver.ResolveShape(_excelShapes.Item(i));
					_items.Add(shape);
				}
			}
		}

        public override bool Equals(IShapes obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Shapes charts = (Shapes)obj;
            return _excelShapes.Equals(charts._excelShapes);
        }
	}
}
