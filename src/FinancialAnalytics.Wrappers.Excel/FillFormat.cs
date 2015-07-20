using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class FillFormat : ExcelEntityWrapper<IFillFormat>, IFillFormat
	{
		protected Microsoft.Office.Interop.Excel.FillFormat _excelFillFormat;

		public FillFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.FillFormat fillFormat)
			: base(entityResolver)
		{
			if (fillFormat == null)
				throw new ArgumentNullException("fillFormat");
			_excelFillFormat = fillFormat;
		}

		public bool Visible
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return MsoTriStateToBoolConverter.Convert(_excelFillFormat.Visible);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelFillFormat.Visible = MsoTriStateToBoolConverter.ConvertBack(value);
				} 
			}
		}

		public IColorFormat ForeColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveColorFormat(_excelFillFormat.ForeColor);
				}
			}
		}

		public float Transparency
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelFillFormat.Transparency;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelFillFormat.Transparency = value;
				}
			}
		}

		public override bool Equals(IFillFormat obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			FillFormat fillFormat = (FillFormat)obj;
			return _excelFillFormat.Equals(fillFormat._excelFillFormat);
		}		
		
		#region Disposable pattern

		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelFillFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion
	}
}
