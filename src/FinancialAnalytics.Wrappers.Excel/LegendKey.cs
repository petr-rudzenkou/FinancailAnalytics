using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
	internal class LegendKey : ExcelEntityWrapper<ILegendKey>, ILegendKey
	{
		private Microsoft.Office.Interop.Excel.LegendKey _excelLegendKey;

        public LegendKey(
			ExcelEntityResolver entityResolver,
			Microsoft.Office.Interop.Excel.LegendKey legendKey)
            : base(entityResolver)
        {
			if (legendKey == null)
				throw new ArgumentNullException("_excelLegendKey");
			_excelLegendKey = legendKey;
        }

		~LegendKey()
        {
            try
            {
                Dispose(false);
            }
            catch (Exception)
            {
            }
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
				ComObjectsFinalizer.ReleaseComObject(_excelLegendKey);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public IInterior Interior
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveInterior(_excelLegendKey.Interior);
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelLegendKey.Border);
				}
			}
		}

		public IChartFillFormat Fill
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFillFormat(_excelLegendKey.Fill);
				}
			}
		}

		public int MarkerForegroundColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelLegendKey.MarkerForegroundColor;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLegendKey.MarkerForegroundColor = value;
				}
			}
		}

		public int MarkerBackgroundColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelLegendKey.MarkerBackgroundColor;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLegendKey.MarkerBackgroundColor = value;
				}
			}
		}

		public ColorIndex MarkerForegroundColorIndex
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlColorIndexToColorIndexConverter.Convert(_excelLegendKey.MarkerForegroundColorIndex);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLegendKey.MarkerForegroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
				}
			}
		}

		public ColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlColorIndexToColorIndexConverter.Convert(_excelLegendKey.MarkerBackgroundColorIndex);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLegendKey.MarkerBackgroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
				}
			}
		}

		public MarkerStyle MarkerStyle
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlMarkerStyleToMarkerStyleConverter.Convert(_excelLegendKey.MarkerStyle);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelLegendKey.MarkerStyle = XlMarkerStyleToMarkerStyleConverter.Convert(value);
				}
			}
		}

		public override bool Equals(ILegendKey obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			LegendKey legendKey = (LegendKey)obj;
			return _excelLegendKey.Equals(legendKey._excelLegendKey);
		}
	}
}
