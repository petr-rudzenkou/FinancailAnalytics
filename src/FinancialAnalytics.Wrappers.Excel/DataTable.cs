using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class DataTable : ExcelEntityWrapper<IDataTable>, IDataTable
	{
		protected Microsoft.Office.Interop.Excel.DataTable _excelDataTable;

		public DataTable(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DataTable dataTable)
			: base(entityResolver)
		{
			if (dataTable == null)
				throw new ArgumentNullException("dataTable");
			_excelDataTable = dataTable;
		}

		#region Disposable pattern

		private bool disposed = false;
		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelDataTable);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(IDataTable other)
		{
			if (other == null || GetType() != other.GetType())
			{
				return false;
			}

			DataTable dataTable = (DataTable)other;
			return _excelDataTable.Equals(dataTable._excelDataTable);
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelDataTable.Border);
				}
			}
		}

		public IFont Font
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveFont(_excelDataTable.Font);
				}
			}
		}

        public IChartFormat Format
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFormat(_excelDataTable.Format);
                }
            }
        }
	}
}
