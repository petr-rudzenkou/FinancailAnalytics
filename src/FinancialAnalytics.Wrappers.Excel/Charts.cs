using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Charts : EntitiesCollectionWrapperBase<ICharts, IChart>, ICharts
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Sheets _excelSheets;

        public Charts(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Sheets sheets)
        {
            if (sheets == null)
                throw new ArgumentNullException("sheets");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _entityResolver = entityResolver;
            _excelSheets = sheets;
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
				ComObjectsFinalizer.ReleaseComObject(_excelSheets);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                for (int i = 1; i <= _excelSheets.Count; i++)
                {
					dynamic excelChart = _excelSheets[i];
					if (excelChart is Microsoft.Office.Interop.Excel.Chart)
					{
						IChart chart =
							_entityResolver.ResolveChart(excelChart as Microsoft.Office.Interop.Excel.Chart);
						_items.Add(chart);
					}
					else
					{
						Marshal.ReleaseComObject(excelChart);
					}
                }
            }
        }

        public IChart AddLast()
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Chart newExcelChart =
                    (Microsoft.Office.Interop.Excel.Chart)
                    _excelSheets.Add(Type.Missing, _excelSheets[_excelSheets.Count],
                                     Type.Missing, Microsoft.Office.Interop.Excel.XlSheetType.xlChart);
                IChart chart = _entityResolver.ResolveChart(newExcelChart);
                _items.Add(chart);
                return chart;
            }
        }

        public IChart AddFirst()
        {
            using (new EnUsCultureInvoker())
            {
                Microsoft.Office.Interop.Excel.Chart newExcelChart =
                    (Microsoft.Office.Interop.Excel.Chart) _excelSheets.Add(_excelSheets[1], Type.Missing,
                                                                            Type.Missing,
                                                                            Microsoft.Office.Interop.Excel.XlSheetType.
                                                                                xlChart);
                IChart chart = _entityResolver.ResolveChart(newExcelChart);
                _items.Add(chart);
                return chart;
            }
        }

        public override bool Equals(ICharts obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Charts charts = (Charts)obj;
            return _excelSheets.Equals(charts._excelSheets);
        }
    }
}
