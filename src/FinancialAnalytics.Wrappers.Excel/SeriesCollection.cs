using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class SeriesCollection : EntitiesCollectionWrapperBase<ISeriesCollection, ISeries>, ISeriesCollection
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.SeriesCollection _excelSeriesColection;

        public SeriesCollection(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection)
        {
            if (seriesCollection == null)
                throw new ArgumentNullException("seriesCollection");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelSeriesColection = seriesCollection;
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
				ComObjectsFinalizer.ReleaseComObject(_excelSeriesColection);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public ISeries Add(IRange source, RowCol rowCol)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.XlRowCol xlRowCol = XlRowColToPlotWayConverter.ConvertBack(rowCol);
                Microsoft.Office.Interop.Excel.Series nativeSeries = _excelSeriesColection.Add((Microsoft.Office.Interop.Excel.Range)source.RangeObject,
                                                                                         xlRowCol, Type.Missing, Type.Missing, Type.Missing);
                ISeries series = _entityResolver.ResolveSeries(nativeSeries);
                _items.Add(series);
                return series;
            }
        }

        public ISeries Add(IRange source, RowCol rowCol, bool seriesLabels, bool categoryLabels, bool replace)
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.XlRowCol xlRowCol = XlRowColToPlotWayConverter.ConvertBack(rowCol);
                Microsoft.Office.Interop.Excel.Series nativeSeries = _excelSeriesColection.Add((Microsoft.Office.Interop.Excel.Range)source.RangeObject,
                                                                                         xlRowCol, seriesLabels, categoryLabels, replace);
                ISeries series = _entityResolver.ResolveSeries(nativeSeries);
                _items.Add(series);
                return series;
            }
        }

		public ISeries NewSeries()
		{
			using (new EnUsCultureInvoker())
			{
				Microsoft.Office.Interop.Excel.Series nativeSeries = _excelSeriesColection.NewSeries();
				ISeries series = _entityResolver.ResolveSeries(nativeSeries);
				_items.Add(series);
				return series;
			}
		}
		
		private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelSeriesColection.Count; i++)
                {
                    try
                    {
                        ISeries series = _entityResolver.ResolveSeries(_excelSeriesColection.Item(i));
                        _items.Add(series);
                    }
                    catch
                    {
                    }
                }
            }
        }

        public override bool Equals(ISeriesCollection obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            SeriesCollection seriesCollection = (SeriesCollection)obj;
            return _excelSeriesColection.Equals(seriesCollection._excelSeriesColection);
        }
    }
}
