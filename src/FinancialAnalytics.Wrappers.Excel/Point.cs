using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using IBorder = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorder;
using IChartFillFormat = FinancialAnalytics.Wrappers.Excel.Interfaces.IChartFillFormat;
using IDataLabel = FinancialAnalytics.Wrappers.Excel.Interfaces.IDataLabel;
using IInterior = FinancialAnalytics.Wrappers.Excel.Interfaces.IInterior;
using IPoint = FinancialAnalytics.Wrappers.Excel.Interfaces.IPoint;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Point : ExcelEntityWrapper<IPoint>, IPoint
    {
        protected Microsoft.Office.Interop.Excel.Point _excelPoint;

        public Point(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Point point)
            : base(entityResolver)
        {
            if (point == null)
                throw new ArgumentNullException("point");
            _excelPoint = point;
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
                ComObjectsFinalizer.ReleaseComObject(_excelPoint);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public bool HasDataLabel
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPoint.HasDataLabel;
                }
            } 
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPoint.HasDataLabel = value;
                }
            } 
        }

        public IDataLabel DataLabel
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveDataLabel(_excelPoint.DataLabel);
                }
            }
        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelPoint.Border);
                }
            } 
        }
        
        public IInterior Interior
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveInterior(_excelPoint.Interior);
                }
            }
        }

        public IChartFillFormat Fill
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartFillFormat(_excelPoint.Fill);
                }
            }
        }

        public object Select()
        {
            using (new EnUsCultureInvoker())
            {
                return _excelPoint.Select();
            }
        }

		public int MarkerForegroundColor
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPoint.MarkerForegroundColor;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPoint.MarkerForegroundColor = value;
				}
			}
		}

        public int MarkerBackgroundColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelPoint.MarkerBackgroundColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelPoint.MarkerBackgroundColor = value;
                }
            }
        }

		public ColorIndex MarkerForegroundColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
					return XlColorIndexToColorIndexConverter.Convert(_excelPoint.MarkerForegroundColorIndex);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
					_excelPoint.MarkerForegroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
                }
            }
        }

		public ColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlColorIndexToColorIndexConverter.Convert(_excelPoint.MarkerBackgroundColorIndex);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPoint.MarkerBackgroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
				}
			}
		}

    	public MarkerStyle MarkerStyle
    	{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return XlMarkerStyleToMarkerStyleConverter.Convert(_excelPoint.MarkerStyle);
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPoint.MarkerStyle = XlMarkerStyleToMarkerStyleConverter.Convert(value);
				}
			}    		
    	}

    	public int MarkerSize
    	{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return _excelPoint.MarkerSize;
				}
			}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelPoint.MarkerSize = value;
				}
			}     		
    	}

        public object Parent
        {
            get
            {
                object parentPbject = null;

                if (_excelPoint.Parent is Microsoft.Office.Interop.Excel.Series)
                {
                    parentPbject = EntityResolver.ResolveSeries((Microsoft.Office.Interop.Excel.Series)_excelPoint.Parent);
                }

                return parentPbject;
            }
        }

        public override bool Equals(IPoint obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Point point = (Point)obj;
            return _excelPoint.Equals(point);
        }
    }
}
