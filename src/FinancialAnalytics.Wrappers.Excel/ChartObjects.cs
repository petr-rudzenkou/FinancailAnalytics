using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ChartObjects : EntitiesCollectionWrapperBase<IChartObjects, IChartObject>, IChartObjects
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.ChartObjects _excelChartObjects;

        public ChartObjects(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartObjects chartObjects)
        {
            if (chartObjects == null)
                throw new ArgumentNullException("chartObjects");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelChartObjects = chartObjects;
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
				ComObjectsFinalizer.ReleaseComObject(_excelChartObjects);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public IChartObject Add(double left, double top, double width, double height)
        {
            using (new EnUsCultureInvoker())
            {
                return _entityResolver.ResolveChartObject(_excelChartObjects.Add(left, top, width, height));
            }
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelChartObjects.Count; i++)
                {
                    IChartObject chartObject =
                        _entityResolver.ResolveChartObject(
                            _excelChartObjects.Item(i) as Microsoft.Office.Interop.Excel.ChartObject);
                    _items.Add(chartObject);
                }
            }
        }
        
        public override bool Equals(IChartObjects obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartObjects chartObjects = (ChartObjects)obj;
            return _excelChartObjects.Equals(chartObjects._excelChartObjects);
        }
    }
}
