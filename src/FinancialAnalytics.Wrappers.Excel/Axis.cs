using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IAxis = FinancialAnalytics.Wrappers.Excel.Interfaces.IAxis;
using IAxisTitle = FinancialAnalytics.Wrappers.Excel.Interfaces.IAxisTitle;
using ITickLabels = FinancialAnalytics.Wrappers.Excel.Interfaces.ITickLabels;
using IGridlines = FinancialAnalytics.Wrappers.Excel.Interfaces.IGridlines;

namespace FinancialAnalytics.Wrappers.Excel
{

    internal class Axis : ExcelEntityWrapper<IAxis>, IAxis
    {
        protected Microsoft.Office.Interop.Excel.Axis _excelAxis;

        public Axis(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Axis axis)
            : base(entityResolver)
        {
            if (axis == null)
                throw new ArgumentNullException("axis");
            _excelAxis = axis;
        }

        public IAxisTitle AxisTitle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveAxisTitle(_excelAxis.AxisTitle);
                }
            }
        }

        public TickLabelPosition TickLabelPosition
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlTickLabelPositionToTickLabelPositionConverter.Convert(_excelAxis.TickLabelPosition);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.TickLabelPosition = XlTickLabelPositionToTickLabelPositionConverter.ConvertBack(value);
                }
            }
        }

        public ITickLabels TickLabels
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveTickLabels(_excelAxis.TickLabels);
                }
            }
        }

        public bool HasMajorGridlines
        {
            get 
            {  
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.HasMajorGridlines;
                } 
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.HasMajorGridlines = value;
                }                 
            }
        }

        public bool HasMinorGridlines
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.HasMinorGridlines;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.HasMinorGridlines = value;
                }
            }
        }

        public bool ReversePlotOrder
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.ReversePlotOrder;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.ReversePlotOrder = value;
                }
            }
        }

        public bool HasTitle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.HasTitle;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.HasTitle = value;
                }
            }
        }

        public ScaleType ScaleType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlScaleTypeToScaleTypeConverter.Convert(_excelAxis.ScaleType);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.ScaleType = XlScaleTypeToScaleTypeConverter.ConvertBack(value);
                }
            }
        }

        public CategoryType CategoryType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlCategoryTypeToCateroryTypeConverter.Convert(_excelAxis.CategoryType);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.CategoryType = XlCategoryTypeToCateroryTypeConverter.ConvertBack(value);
                }
            }
        }

        public AxisCrosses Crosses
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlAxisCrossesToAxisCrossesConverter.Convert(_excelAxis.Crosses);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.Crosses = XlAxisCrossesToAxisCrossesConverter.Convert(value);
                }
            }
        }

        public double CrossesAt
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.CrossesAt;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.CrossesAt = value;
                }
            }
        }

        public double MinimumScale
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.MinimumScale;
                }
            }
            set
            {
                MinimumScaleIsAuto = false;
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.MinimumScale = value;
                }
            }
        }

        public double MaximumScale
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.MaximumScale;
                }
            }
            set
            {
                MaximumScaleIsAuto = false;
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.MaximumScale = value;
                }
            }
        }

        public bool MaximumScaleIsAuto
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.MaximumScaleIsAuto;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.MaximumScaleIsAuto = value;
                }
            }
        }

        public bool MinimumScaleIsAuto
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelAxis.MinimumScaleIsAuto;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelAxis.MinimumScaleIsAuto = value;
                }
            }
        }

        public Interfaces.IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_excelAxis.Border);
                }
            }
        }

    	public DisplayUnit DisplayUnit
    	{
    		get
    		{
    			using (new EnUsCultureInvoker())
    			{
					return XlDisplayUnitToDisplayUnitConverter.Convert(_excelAxis.DisplayUnit);
    			}
    		}
			set
			{
				using (new EnUsCultureInvoker())
				{
					if (value != DisplayUnit.IsNotSet)
					{
						_excelAxis.DisplayUnit = XlDisplayUnitToDisplayUnitConverter.ConvertBack(value);
					}
				}
			}
    	}

		public double DisplayUnitCustom
		{
    		get
    		{
    			using (new EnUsCultureInvoker())
    			{
    				return _excelAxis.DisplayUnitCustom;
    			}
    		}
			set
			{
				using (new EnUsCultureInvoker())
				{
					_excelAxis.DisplayUnitCustom = value;
				}
			}
		}

		public IGridlines MajorGridlines
    	{
    		get
    		{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveGridlines(_excelAxis.MajorGridlines);
				}    			
    		}
    	}


		public IGridlines MinorGridlines
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveGridlines(_excelAxis.MinorGridlines);
				}
			}
		}

        public override bool Equals(IAxis obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Axis axis = (Axis)obj;
            return _excelAxis.Equals(axis._excelAxis);
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
                ComObjectsFinalizer.ReleaseComObject(_excelAxis);
                disposed = true;
            }
            base.Dispose(disposing);
        }
        #endregion



		public Interfaces.IChartFormat Format
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveChartFormat(_excelAxis.Format);
				}
			}
		}
	}
}
