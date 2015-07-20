using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class LineFormat : ExcelEntityWrapper<ILineFormat>, ILineFormat
	{
		protected Microsoft.Office.Interop.Excel.LineFormat _excelLineFormat;

		public LineFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.LineFormat lineFormat)
			: base(entityResolver)
		{
			if (lineFormat == null)
				throw new ArgumentNullException("lineFormat");
			_excelLineFormat = lineFormat;
		}

		public bool Visible
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return MsoTriStateToBoolConverter.Convert(_excelLineFormat.Visible);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLineFormat.Visible = MsoTriStateToBoolConverter.ConvertBack(value);
				} 
			}
		}

		public float Weight
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelLineFormat.Weight;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLineFormat.Weight = value;
				}
			}
		}

		public IColorFormat ForeColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveColorFormat(_excelLineFormat.ForeColor);
				}
			}
		}

		public IColorFormat BackColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveColorFormat(_excelLineFormat.BackColor);
				}
			}
		}

		public float Transparency
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelLineFormat.Transparency;	
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLineFormat.Transparency = value; 
				}
			}
		}

		public override bool Equals(ILineFormat obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			LineFormat lineFormat = (LineFormat)obj;
			return _excelLineFormat.Equals(lineFormat._excelLineFormat);
		}		
		
		#region Disposable pattern

		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelLineFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion
	}
}
