using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class Theme : ExcelEntityWrapper<ITheme>, ITheme
    {
		private Microsoft.Office.Core.OfficeTheme _excelTheme;

        public Theme(ExcelEntityResolver entityResolver, Microsoft.Office.Core.OfficeTheme theme)
            : base(entityResolver)
        {
            if (theme == null)
                throw new ArgumentNullException("theme");
            _excelTheme = theme;
        }

        public IThemeColorScheme ThemeColorScheme
        {
            get
            {
				return EntityResolver.ResolveThemeColorScheme(_excelTheme.ThemeColorScheme);
            }
        }

        public IThemeFontScheme ThemeFontScheme
        {
            get
            {
                return EntityResolver.ResolveThemeFontScheme(_excelTheme.ThemeFontScheme);
            }
        }

        public override bool Equals(ITheme obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            Theme theme = (Theme)obj;
            return _excelTheme.Equals(theme._excelTheme);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelTheme);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
