using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class Axes : EntitiesCollectionWrapperBase<IAxes, IAxis>, IAxes
    {
        protected ExcelEntityResolver _entityResolver;
        Microsoft.Office.Interop.Excel.Axes _excelAxes;

        public Axes(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Axes axes)
        {
            if (axes == null)
                throw new ArgumentNullException("axes");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelAxes = axes;
            _entityResolver = entityResolver;
            InitializeCollection();
        }

        ~Axes()
        {
            try
            {
                Dispose();
            }
            catch (Exception)
            {
            }
        }

        private void InitializeCollection()
        {
            using (new EnUsCultureInvoker())
            {
                foreach(Microsoft.Office.Interop.Excel.Axis excelAxis in _excelAxes)
                {
                    IAxis axis = _entityResolver.ResolveAxis(excelAxis);
                    _items.Add(axis);
                }
            }
        }

        public override bool Equals(IAxes obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Axes axes = (Axes)obj;
            return _excelAxes.Equals(axes._excelAxes);
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
                ComObjectsFinalizer.ReleaseComObject(_excelAxes);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        public IAxis Item(AxisType type)
        {
            Microsoft.Office.Interop.Excel.XlAxisType xlAxisType = XlAxisTypeToAxisTypeConverter.ConvertBack(type);
            return _entityResolver.ResolveAxis(_excelAxes.Item(xlAxisType));
        }

    }
}
