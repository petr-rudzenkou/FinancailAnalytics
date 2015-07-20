using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{
    class Walls : ExcelEntityWrapper<IWalls>, IWalls
    {
        private Microsoft.Office.Interop.Excel.Walls _excelWalls;

        public Walls(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Walls walls)
            : base(entityResolver)
        {
            if (walls == null)
                throw new ArgumentNullException("walls");
            _excelWalls = walls;
        }

        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelWalls.Interior);
                }
            }
        }

        public IChartFormat Format
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFormat(_excelWalls.Format);
                }
            }
        }


        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelWalls.Border);
                }
            }
        }

        public override bool Equals(IWalls obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Walls walls = (Walls)obj;
            return _excelWalls.Equals(walls._excelWalls);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelWalls);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion


    }
}
