using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ChartFillFormat : ExcelEntityWrapper<IChartFillFormat>, IChartFillFormat
    {
        protected Microsoft.Office.Interop.Excel.ChartFillFormat _excelChartFillFormat;

        public ChartFillFormat(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.ChartFillFormat chartFillFormat)
            : base(entityResolver)
        {
            if (chartFillFormat == null)
                throw new ArgumentNullException("chartFillFormat");
            _excelChartFillFormat = chartFillFormat;
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
				ComObjectsFinalizer.ReleaseComObject(_excelChartFillFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public bool Visible
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoTriStateToBoolConverter.Convert(_excelChartFillFormat.Visible);
                }
            }            
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelChartFillFormat.Visible = MsoTriStateToBoolConverter.ConvertBack(value);
                }                
            }
        }

        public FillType Type
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoFillTypeToFillTypeConverter.Convert(_excelChartFillFormat.Type);
                }                
            }
        }

        public IChartColorFormat ForeColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartColorFormat(_excelChartFillFormat.ForeColor);
                }
            }
        }

        public IChartColorFormat BackColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return EntityResolver.ResolveChartColorFormat(_excelChartFillFormat.BackColor);
                }
            }
        }

        public int GradientVariant
        { 
            get
            {
                 using (new EnUsCultureInvoker())
                 {
                     return _excelChartFillFormat.GradientVariant;
                 }
            } 
        }

        public float GradientDegree
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelChartFillFormat.GradientDegree;
                }
            }
        }

        public PatternType Pattern
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    {
                        return MsoPatternTypeToPatternTypeConverter.Convert(_excelChartFillFormat.Pattern);
                    }
                }
            }
        }

        public GradientColorType GradientColorType
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoGradientColorTypeToGradientColorType.Convert(_excelChartFillFormat.GradientColorType);
                }
            }
        }

        public GradientStyle GradientStyle
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoGradientStyleToGradientStyleConverter.Convert(_excelChartFillFormat.GradientStyle);
                }
            }
        }

        public PresetTexture PresetTexture
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return MsoPresetTextureToPresetTextureConverter.Convert(_excelChartFillFormat.PresetTexture);
                }                
            }
        }

        public PresetGradientType PresetGradientType
        {
            get
            {
               using (new EnUsCultureInvoker())
               {
                   return MsoPresetGradientTypeToPresetGradientTypeCoverter.Convert(_excelChartFillFormat.PresetGradientType);
               } 
            }
        }

        public void Solid()
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.Solid();
            }
        }

        public void OneColorGradient(GradientStyle style, int variant, float degree)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.OneColorGradient(MsoGradientStyleToGradientStyleConverter.ConvertBack(style), variant,  degree);
            }
        }

        public void TwoColorGradient(GradientStyle style, int variant, float degree)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.OneColorGradient(MsoGradientStyleToGradientStyleConverter.ConvertBack(style),
                                                       variant, degree);
            }
        }

        public void PresetGradient(GradientStyle style, int variant, PresetGradientType presetGradientType)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.PresetGradient(MsoGradientStyleToGradientStyleConverter.ConvertBack(style),
                                                       variant, MsoPresetGradientTypeToPresetGradientTypeCoverter.ConvertBack(presetGradientType));
            }            
        }

        public void PresetTextured(PresetTexture presetTexture)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.PresetTextured(MsoPresetTextureToPresetTextureConverter.ConvertBack(presetTexture));
            }
        }


        public void Patterned(PatternType pattern)
        {
            using (new EnUsCultureInvoker())
            {
                _excelChartFillFormat.Patterned(MsoPatternTypeToPatternTypeConverter.ConvertBack(pattern));
            }
        }

        public override bool Equals(IChartFillFormat obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ChartFillFormat chartFillFormat = (ChartFillFormat)obj;
            return _excelChartFillFormat.Equals(chartFillFormat._excelChartFillFormat);
        }
    }
}
