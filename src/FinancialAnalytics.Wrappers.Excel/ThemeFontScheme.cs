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
	class ThemeFontScheme : ExcelEntityWrapper<IThemeFontScheme>, IThemeFontScheme
	{
		private Microsoft.Office.Core.ThemeFontScheme _excelThemeFontScheme;

		public ThemeFontScheme(ExcelEntityResolver entityResolver, Microsoft.Office.Core.ThemeFontScheme themeFontScheme)
			: base(entityResolver)
		{
			if (themeFontScheme == null)
				throw new ArgumentNullException("themeFontScheme");
			_excelThemeFontScheme = themeFontScheme;
		}

		public void Save(string fileName)
		{
			_excelThemeFontScheme.Save(fileName);
		}

		public void Load(string fileName)
		{
			_excelThemeFontScheme.Load(fileName);
		}

		public override bool Equals(IThemeFontScheme obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ThemeFontScheme themeFontScheme = (ThemeFontScheme)obj;
			return _excelThemeFontScheme.Equals(themeFontScheme._excelThemeFontScheme);
		}

		#region Disposable pattern

		private bool disposed = false;

		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelThemeFontScheme);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

	}
}
