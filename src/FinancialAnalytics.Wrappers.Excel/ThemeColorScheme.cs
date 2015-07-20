using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    class ThemeColorScheme : ExcelEntityWrapper<IThemeColorScheme>, IThemeColorScheme
    {
		private Microsoft.Office.Core.ThemeColorScheme _excelThemeColorScheme;

		public ThemeColorScheme(ExcelEntityResolver entityResolver, Microsoft.Office.Core.ThemeColorScheme themeColorScheme)
            : base(entityResolver)
        {
            if (themeColorScheme == null)
                throw new ArgumentNullException("themeColorScheme");
            _excelThemeColorScheme = themeColorScheme;
        }

        public void Save(string fileName)
        {
			_excelThemeColorScheme.Save(fileName);
        }

        public void Load(string fileName)
        {
			_excelThemeColorScheme.Load(fileName);
        }

        public override bool Equals(IThemeColorScheme obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            ThemeColorScheme themeColorScheme = (ThemeColorScheme)obj;
            return _excelThemeColorScheme.Equals(themeColorScheme._excelThemeColorScheme);
        }

        #region Disposable pattern

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                ComObjectsFinalizer.ReleaseComObject(_excelThemeColorScheme);
                disposed = true;
            }
            base.Dispose(disposing);
        }

        #endregion

    }
}
