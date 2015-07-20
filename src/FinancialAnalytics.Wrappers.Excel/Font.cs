using System;
using System.Drawing;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using IFont = FinancialAnalytics.Wrappers.Excel.Interfaces.IFont;

namespace FinancialAnalytics.Wrappers.Excel
{
    public class Font : ExcelEntityWrapper<IFont>, IFont
    {
        protected Microsoft.Office.Interop.Excel.Font _excelFont;
        private static readonly Object _locker = new object();

        public Font(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.Font font)
            : base(entityResolver)
        {
            if (font == null)
                throw new ArgumentNullException("font");
            _excelFont = font;
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
				ComObjectsFinalizer.ReleaseComObject(_excelFont);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion

        public object Size
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Size;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Size = value;
                }
            }
        }

        public Color Color
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
					if (_excelFont == null || _excelFont.Color is System.DBNull)
 					{
						return Color.Empty;
					}
					else
					{
						if (_excelFont.Color is int)
						{
							return ColorTranslator.FromOle((int)_excelFont.Color);
						}
						return ColorTranslator.FromOle((int)(double)_excelFont.Color);
					}
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Color = ColorTranslator.ToOle(value);
                }
            }
        }

        public object ColorIndex
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.ColorIndex;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.ColorIndex = value;
                }
            }
        }

        public object TintAndShade
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.TintAndShade;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.TintAndShade = value;
                }
            }
        }

        public object ThemeColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.ThemeColor;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.ThemeColor = value;
                }
            }
        }


        /// <summary>
        /// Use this color for Chart.Font.Color at 2007/2010
        /// </summary>
        public Color ChartColor
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                   return ColorTranslator.FromOle((int)_excelFont.Color);
                    
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Color = ColorTranslator.ToOle(value);
                }
            }
        }

        public string Name
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return (string)_excelFont.Name;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Name = value;
                }
            }
        }

        public UnderlineStyle Underline
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return XlUnderlineStyleToUnderlineStyleConverter.Convert((XlUnderlineStyle)_excelFont.Underline);
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Underline = (object)XlUnderlineStyleToUnderlineStyleConverter.ConvertBack(value);
                }
            }
        }

        public object UnderlineAsObject
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Underline;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Underline = value;
                }
            }
        }

        public bool Bold
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Bold is bool ? (bool)_excelFont.Bold : false;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    //lock(_locker)
                    //{
                        _excelFont.Bold = value;
                    //}
                }
            }
        }

        public bool Italic
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Italic is bool && (bool)_excelFont.Italic;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Italic = value;
                }
            }
        }

        public bool Strikethrough
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Strikethrough is bool && (bool)_excelFont.Strikethrough;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Strikethrough = value;
                }
            }
        }

        public bool Subscript
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Subscript is bool && (bool)_excelFont.Subscript;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Subscript = value;
                }
            }            
        }

        public bool Superscript
        {
            get
            {
                using (new EnUsCultureInvoker())
                {
                    return _excelFont.Superscript is bool && (bool)_excelFont.Superscript;
                }
            }
            set
            {
                using (new EnUsCultureInvoker())
                {
                    _excelFont.Superscript = value;
                }
            }             
        }

        public override bool Equals(IFont obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Font font = (Font)obj;
            return _excelFont.Equals(font._excelFont);
        }
    }
}
