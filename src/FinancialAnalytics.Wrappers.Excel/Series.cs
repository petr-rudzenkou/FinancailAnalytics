using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Converters;

namespace FinancialAnalytics.Wrappers.Excel
{
    /// <remarks>
    /// Pattern "values = (Array) ((object)_underlyingObject.XValues);"  fixes 
    /// http://stackoverflow.com/questions/4807968/net-4-0-excel-interop-issues-with-dynamic-collections
    /// </remarks>
    internal class Series : ExcelEntityWrapper<ISeries>, ISeries
    {
        protected Microsoft.Office.Interop.Excel.Series _underlyingObject;
        private bool disposed = false;

        public Series(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Series series)
            : base(entityResolver)
        {
            if (series == null)
                throw new ArgumentNullException("series");
            _underlyingObject = series;
        }

        public Array Values
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return GetValues();
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    SetValues(value);
                }
            }
        }

        public Array XValues
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return GetXValues();
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    SetXValues(value);
                }
            }
        }

        public object NativeValues
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.Values;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.Values = value;
                }
            }
        }

        public object NativeXValues
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.XValues;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.XValues = value;
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.Name;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.Name = value;
                }
            }
        }

        public string Formula
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.Formula;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.Formula = value;
                }
            }
        }

        public string FormulaLocal
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.FormulaLocal;
                }
            }
        }

        public bool HasDataLabels
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.HasDataLabels;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.HasDataLabels = value;
                }
            }
        }

        protected virtual Array GetValues()
        {
            using (new EnUsCultureInvoker())
            {
                Array values;
                try
                {
                    values = (Array)((object)_underlyingObject.Values);
                }
                catch (Exception exc)
                {
                    bool rethrow = ExceptionHandler.HandleException(exc);
                    if (rethrow)
                        throw;
                    values = new string[((Array)_underlyingObject.Values).Length];
                }
                return values;
            }
        }

        public IPoints Points(object index)
        {
            using (new EnUsCultureInvoker())
            {
                return EntityResolver.ResolvePoints(_underlyingObject.Points(index) as Microsoft.Office.Interop.Excel.Points);
            }
        }

        public object Select()
        {
            using (new EnUsCultureInvoker())
            {
                return _underlyingObject.Select();
            }
        }

        public void Delete()
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Delete();
            }
        }

        public void Copy()
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Copy();
            }
        }

        public void Paste()
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Paste();
            }
        }

        protected virtual void SetValues(Array values)
        {
            using (new EnUsCultureInvoker())
            {
                _underlyingObject.Values = values;
            }
        }

        public ChartType ChartType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlChartTypeToChartTypeConverter.Convert(_underlyingObject.ChartType);
                }
            }
        }

        protected virtual Array GetXValues()
        {
            using (new EnUsCultureInvoker())
            {
                Array values;
                try
                {
                    values = (Array)((object)_underlyingObject.XValues);
                }
                catch (Exception exc)
                {
                    bool rethrow = ExceptionHandler.HandleException(exc);
                    if (rethrow)
                        throw;
                    values = new string[((Array)_underlyingObject.XValues).Length];
                }
                return values;
            }
        }

        protected virtual void SetXValues(Array values)
        {

        }

        public IBorder Border
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveBorder(_underlyingObject.Border);
                }
            }
        }

        public IChartFillFormat Fill
        {
            get { return EntityResolver.ResolveChartFillFormat(_underlyingObject.Fill); }
        }

        public IInterior Interior
        {
            get { return EntityResolver.ResolveInterior(_underlyingObject.Interior); }
        }

        public int MarkerForegroundColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.MarkerForegroundColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerForegroundColor = value;
                }
            }
        }

        public int MarkerBackgroundColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.MarkerBackgroundColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerBackgroundColor = value;
                }
            }
        }

        public ColorIndex MarkerForegroundColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlColorIndexToColorIndexConverter.Convert(_underlyingObject.MarkerForegroundColorIndex);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerForegroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
                }
            }
        }

        public ColorIndex MarkerBackgroundColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlColorIndexToColorIndexConverter.Convert(_underlyingObject.MarkerBackgroundColorIndex);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerBackgroundColorIndex = XlColorIndexToColorIndexConverter.ConvertBack(value);
                }
            }
        }

        public MarkerStyle MarkerStyle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlMarkerStyleToMarkerStyleConverter.Convert(_underlyingObject.MarkerStyle);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerStyle = XlMarkerStyleToMarkerStyleConverter.Convert(value);
                }
            }
        }

        public int MarkerSize
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _underlyingObject.MarkerSize;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _underlyingObject.MarkerSize = value;
                }
            }
        }

        public bool HasUndefinedType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return (int)_underlyingObject.ChartType == -4111;
                }
            }
        }

        public override bool Equals(ISeries obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            var chartTitle = (Series)obj;
            return _underlyingObject.Equals(chartTitle._underlyingObject);
        }


        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_underlyingObject);
                disposed = true;
            }
            base.Dispose(disposing);
        }
    }
}
