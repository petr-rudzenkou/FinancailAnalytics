using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ColorFormat : ExcelEntityWrapper<IColorFormat>, IColorFormat
    {
        protected Microsoft.Office.Interop.Excel.ColorFormat _excelColorFormat;

        public ColorFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ColorFormat colorFormat)
            : base(entityResolver)
        {
            if (colorFormat == null)
                throw new ArgumentNullException("colorFormat");
            _excelColorFormat = colorFormat;
        }

        public int RGB
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelColorFormat.RGB;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelColorFormat.RGB = value;
                } 
            }
        }

        public override bool Equals(IColorFormat obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ColorFormat colorFormat = (ColorFormat)obj;
            return _excelColorFormat.Equals(colorFormat._excelColorFormat);
        }	

		public Office.Enums.ColorType Type
		{
			get { return MsoColorTypeToColorTypeConverter.Convert(_excelColorFormat.Type); }
		}	
        
        #region Disposable pattern

		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelColorFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion
	}
}
