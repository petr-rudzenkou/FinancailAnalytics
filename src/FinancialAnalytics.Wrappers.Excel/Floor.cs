using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Floor : ExcelEntityWrapper<IFloor>, IFloor
    {
        private Microsoft.Office.Interop.Excel.Floor _excelFloor;

        public Floor(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Floor floor)
            : base(entityResolver)
        {
            if (floor == null)
                throw new ArgumentNullException("floor");
            _excelFloor = floor;
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelFloor.Interior);
                }
            }
        }

        public IChartFormat Format
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFormat(_excelFloor.Format);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelFloor.Border);
                }
            }
        }

        public override bool Equals(IFloor obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Floor floor = (Floor)obj;
            return _excelFloor.Equals(floor._excelFloor);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelFloor);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
        
    }
}
