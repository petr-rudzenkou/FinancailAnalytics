using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IBorder = FinancialAnalytics.Wrappers.Excel.Interfaces.IBorder;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Border : ExcelEntityWrapper<IBorder>, IBorder
    {
        #region Constants and Fields

        protected Microsoft.Office.Interop.Excel.Border _excelBorder;
        private bool _disposed;

        #endregion

        #region Constructors and Destructors

        public Border(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Border border)
            : base(entityResolver)
        {
            if (border == null)
            {
                throw new ArgumentNullException("border");
            }
            _excelBorder = border;
        }

        #endregion

        #region Properties

        public object Color
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorder.Color;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.Color = value;
                }
            }
        }

        public object ColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorder.ColorIndex;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.ColorIndex = value;
                }
            }
        }


        public object ThemeColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorder.ThemeColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.ThemeColor = value;
                }
            }
        }


        public object TintAndShade
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorder.TintAndShade;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.TintAndShade = value;
                }
            }
        }

        public LineStyle LineStyle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlLineStyleToLineStyleConverter.Convert((XlLineStyle)_excelBorder.LineStyle);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    try
                    {
                        _excelBorder.LineStyle = (int) XlLineStyleToLineStyleConverter.ConvertBack(value);
                    }
                    catch (Exception)
                    {
                        _excelBorder.LineStyle = -4105;
                    }
                }
            }
        }

        public object LineStyleObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelBorder.LineStyle;
                }
            }

            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.LineStyle = value;
                }
            }
        }

        public BorderWeight Weight
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlBorderWeightToBorderWeightConverter.Convert((XlBorderWeight)_excelBorder.Weight);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelBorder.Weight = XlBorderWeightToBorderWeightConverter.ConvertBack(value);
                }
            }
        }

        #endregion

        #region Implemented Interfaces

        #region IEquatable<IBorder>

        public override bool Equals(IBorder obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            var border = (Border)obj;
            return _excelBorder.Equals(border._excelBorder);
        }

        #endregion

        #endregion

        #region Methods

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
                ComObjectsFinalizer.ReleaseComObject(_excelBorder);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}