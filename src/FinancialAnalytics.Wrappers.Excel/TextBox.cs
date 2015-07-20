using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;

namespace FinancialAnalytics.Wrappers.Excel
{
	public class TextBox : ExcelEntityWrapper<ITextBox>, ITextBox
	{
		protected Microsoft.Office.Interop.Excel.TextBox _excelTextBox;

		public TextBox(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.TextBox textBox)
			: base(entityResolver)
		{
			if (textBox == null)
				throw new ArgumentNullException("textBox");
			_excelTextBox = textBox;
		}

		#region Disposable pattern

		private bool disposed = false;
		protected override void Dispose(bool disposing)
		{
			if (!disposed)
			{
				ComObjectsFinalizer.ReleaseComObject(_excelTextBox);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(ITextBox obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}

			TextBox textBox = (TextBox)obj;
			return _excelTextBox.Equals(textBox._excelTextBox);
		}

		public IInterior Interior
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveInterior(_excelTextBox.Interior);
				}
			}
		}

		public IBorder Border
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveBorder(_excelTextBox.Border);
				}
			}
		}

		public IFont Font
		{
			get
			{
				using (new EnUsCultureInvoker())
				{
					return EntityResolver.ResolveFont(_excelTextBox.Font);
				}
			}
		}
	}
}
