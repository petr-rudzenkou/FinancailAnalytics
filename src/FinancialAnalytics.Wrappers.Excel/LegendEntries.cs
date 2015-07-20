using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class LegendEntries : EntitiesCollectionWrapperBase<ILegendEntries, ILegendEntry>, ILegendEntries
	{
		protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.LegendEntries _excelLegendEntries;

		public LegendEntries(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.LegendEntries legendEntries)
        {
			if (legendEntries == null)
                throw new ArgumentNullException("seriesCollection");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
			_excelLegendEntries = legendEntries;
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
				ComObjectsFinalizer.ReleaseComObject(_excelLegendEntries);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

		private void InitializeCollection()
		{
			using (new EnUsCultureInvoker())
			{
				for (int i = 1; i <= _excelLegendEntries.Count; i++)
				{
					try
					{
						ILegendEntry legendEntry = _entityResolver.ResolveLegendEntry(_excelLegendEntries.Item(i));
						_items.Add(legendEntry);
					}
					catch
					{
					}
				}
			}
		}

		public override bool Equals(ILegendEntries obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			LegendEntries legendEntries = (LegendEntries)obj;
			return _excelLegendEntries.Equals(legendEntries._excelLegendEntries);
		}
	}
}
