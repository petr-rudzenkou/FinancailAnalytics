using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IBorder = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorder;
using IChartFillFormat = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartFillFormat;
using IDataLabel = FinancialAnalytics.Wrappers.Excel.Interfaces.IDataLabel;
using IInterior = FinancialAnalytics.Wrappers.Excel.Interfaces.IInterior;
using FinancialAnalytics.Wrappers.Excel.Converters;
using ILegendEntry = FinancialAnalytics.Wrappers.Excel.Interfaces.ILegendEntry;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class LegendEntry : ExcelEntityWrapper<ILegendEntry>, ILegendEntry
    {
        protected Microsoft.Office.Interop.Excel.LegendEntry _excelLegendEntry;

        public LegendEntry(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.LegendEntry legendEntry)
            : base(entityResolver)
        {
            if (legendEntry == null)
                throw new ArgumentNullException("legendEntry");
            _excelLegendEntry = legendEntry;
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
                ComObjectsFinalizer.ReleaseComObject(_excelLegendEntry);
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
                    return EntityResolver.ResolveFont(_excelLegendEntry.Font);
                }
            }
        }

		public Wrappers.Excel.Interfaces.ILegendKey LegendKey
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveLegendKey(_excelLegendEntry.LegendKey);
				}
			}
		}

        public override bool Equals(ILegendEntry obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            LegendEntry legendEntry = (LegendEntry)obj;
            return _excelLegendEntry.Equals(legendEntry);
        }

    }
}
