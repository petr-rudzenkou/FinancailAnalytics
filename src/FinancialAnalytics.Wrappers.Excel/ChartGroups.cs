using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class ChartGroups : EntitiesCollectionWrapperBase<IChartGroups, IChartGroup>, IChartGroups
	{
		protected ExcelEntityResolver _entityResolver;
		protected Microsoft.Office.Interop.Excel.ChartGroups _excelChartGroups;

		public ChartGroups(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartGroups chartGroups)
		{
			if (chartGroups == null)
			{
				throw new ArgumentNullException("chartGroups");
			}
			if (entityResolver == null)
			{
				throw new ArgumentNullException("entityResolver");
			}
			_excelChartGroups = chartGroups;
			_entityResolver = entityResolver;
			InitializeCollection();
		}

		public override bool Equals(IChartGroups obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ChartGroups chartGroups = (ChartGroups)obj;
			return _excelChartGroups.Equals(chartGroups._excelChartGroups);
		}

		private void InitializeCollection()
		{
			using (new EnUsCultureInvoker())
			{
				for (int i = 1; i <= _excelChartGroups.Count; i++)
				{
					IChartGroup chartGroup = _entityResolver.ResolveChartGroup(_excelChartGroups.Item(i));
					_items.Add(chartGroup);
				}
			}
		}

		#region Disposable pattern

		private bool _disposed;
		protected override void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelChartGroups);
				_disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
