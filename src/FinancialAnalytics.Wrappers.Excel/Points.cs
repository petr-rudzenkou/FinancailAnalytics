using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Points : EntitiesCollectionWrapperBase<IPoints, IPoint>, IPoints
    {
        protected ExcelEntityResolver _entityResolver;
        Microsoft.Office.Interop.Excel.Points _excelPoints;

        public Points(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Points points)
        {
            if (points == null)
                throw new ArgumentNullException("points");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelPoints = points;
            _entityResolver = entityResolver;
            InitializeCollection();
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
                ComObjectsFinalizer.ReleaseComObject(_excelPoints);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelPoints.Count; i++)
                {
                    IPoint point = _entityResolver.ResolvePoint(_excelPoints.Item(i));
                    _items.Add(point);
                }
            }
        }

        public override bool Equals(IPoints obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Points points = (Points)obj;
            return _excelPoints.Equals(points);
        }
    }
}
