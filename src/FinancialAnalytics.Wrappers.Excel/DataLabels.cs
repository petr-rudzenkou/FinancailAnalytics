using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IDataLabels = FinancialAnalytics.Wrappers.Excel.Interfaces.IDataLabels;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class DataLabels : ExcelEntityWrapper<IDataLabels>, IDataLabels
    {
        protected Microsoft.Office.Interop.Excel.DataLabels _excelDataLabels;

        public DataLabels(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.DataLabels dataLabels)
            : base(entityResolver)
        {
            if (dataLabels == null)
                throw new ArgumentNullException("dataLabels");
            _excelDataLabels = dataLabels;
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
                ComObjectsFinalizer.ReleaseComObject(_excelDataLabels);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
        
        /// <summary>
        /// Returns a Font object that represents the font of the specified object.
        /// </summary>
        public Wrappers.Excel.Interfaces.IFont Font
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveFont(_excelDataLabels.Font);
                }
            }
        }

        public Interfaces.IChartFormat Format
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFormat(_excelDataLabels.Format);
                }
            }
        }

        public override bool Equals(IDataLabels obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            DataLabels dataLabels = (DataLabels)obj;
            return _excelDataLabels.Equals(dataLabels);
        }


        public IBorder Border
        {
            get 
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelDataLabels.Border);
                }
            }
        }
    }
}
