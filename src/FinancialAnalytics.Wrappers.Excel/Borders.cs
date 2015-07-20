using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel
{


    internal class Borders : EntitiesCollectionWrapperBase<IBorders, IBorder>, IBorders
    {
        protected ExcelEntityResolver _entityResolver;
        protected Microsoft.Office.Interop.Excel.Borders _excelBorders;


        public Borders(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Borders borders)
        {
            if (borders == null)
                throw new ArgumentNullException("borders");
            if (entityResolver == null)
                throw new ArgumentNullException("entityResolver");
            _excelBorders = borders;             
            _entityResolver = entityResolver;

        }

        public object Weight
        {
            get { return _excelBorders.Weight; }
            set
            {
                _excelBorders.Weight = value;
            }
        }
       
		public object Color
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelBorders.Color;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelBorders.Color = value;
				}
			}
		}

		public object ColorIndex
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelBorders.ColorIndex;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelBorders.ColorIndex = value;
				}
			}
		}

        public override bool Equals(IBorders obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Borders borders = (Borders)obj;
            return _excelBorders.Equals(borders._excelBorders);
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
                ComObjectsFinalizer.ReleaseComObject(_excelBorders);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion

        public IBorder this[Enums.BordersIndex bordersIndex]
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _entityResolver.ResolveBorder(_excelBorders[XlBordersIndexToBordersIndexConverter.ConvertBack(bordersIndex)]);
                }
            }
        }

        public object LineStyle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorders.LineStyle;
                }
            }
            set 
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorders.LineStyle = value;
                }
            }
        }
    }
}
