using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class SeriesLines : ExcelEntityWrapper<ISeriesLines>, ISeriesLines
    {
        protected Microsoft.Office.Interop.Excel.SeriesLines _excelSeriesLines;

        public SeriesLines(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.SeriesLines excelSeriesLines)
			: base(entityResolver)
		{
            if (excelSeriesLines == null)
			{
                throw new ArgumentNullException("excelSeriesLines");
			}
            _excelSeriesLines = excelSeriesLines;
		}

        public IBorder Border
        {
            get 
            {
                return EntityResolver.ResolveBorder(_excelSeriesLines.Border);
            }
        }

        public override bool Equals(ISeriesLines obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            SeriesLines seriesLines = (SeriesLines)obj;
            return _excelSeriesLines.Equals(seriesLines._excelSeriesLines);
        }

        #region Disposable pattern

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelSeriesLines);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
        
    }
}
