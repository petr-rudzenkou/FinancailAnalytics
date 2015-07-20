using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using IInterior = FinancialAnalytics.Wrappers.Excel.Interfaces.IInterior;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Interior : ExcelEntityWrapper<IInterior>, IInterior
    {
        protected Microsoft.Office.Interop.Excel.Interior _excelInterior;
        private LateBindingInvoker _invoker;

        public Interior(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Interior interior)
            : base(entityResolver)
        {
            if (interior == null)
                throw new ArgumentNullException("interior");
            _excelInterior = interior;
            _invoker = new LateBindingInvoker(_excelInterior);
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
				ComObjectsFinalizer.ReleaseComObject(_excelInterior);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public object ColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.ColorIndex;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.ColorIndex = value;
                }
            }
        }

        public override bool Equals(IInterior obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Interior interior = (Interior)obj;
            return _excelInterior.Equals(interior._excelInterior);
        }



        public object Pattern
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.Pattern;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.Pattern = value;
                }
            }
        }

		public Pattern GetPattern()
		{
			using (new EnUsCultureInvoker())
			{
				return XlPatternToPatternConverter.Convert((XlPattern)_excelInterior.Pattern);
			}
		}

        public object Color
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.Color;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.Color = value;
                }
            }
        }


        public object ThemeColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.ThemeColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.ThemeColor = value;
                }
            }
        }


        public object TintAndShade
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.TintAndShade;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.TintAndShade = value;
                }
            }
        }


        public object PatternColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.PatternColorIndex;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.PatternColorIndex = value;
                }
            }            
        }

        public object PatternColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelInterior.PatternColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelInterior.PatternColor = value;
                }
            }               
        }

        private const string GradientPropertyName = "Gradient";
        public IGradient Gradient
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    int excelPattern = (int)_excelInterior.Pattern;
                    if (excelPattern == (int)XlPatternToPatternConverter.ConvertBack(Enums.Pattern.PatternLinearGradient))
                    {
                        return EntityResolver.ResolveLinearGradient(_invoker.InvokeGetPropertyValue(GradientPropertyName));
                    }
                    if (excelPattern == (int)XlPatternToPatternConverter.ConvertBack(Enums.Pattern.PatternRectangularGradient))
                    {
                        return EntityResolver.ResolveRectangularGradient(_invoker.InvokeGetPropertyValue(GradientPropertyName));
                    }
                    return null;
                }
            }
        }

    }
}
