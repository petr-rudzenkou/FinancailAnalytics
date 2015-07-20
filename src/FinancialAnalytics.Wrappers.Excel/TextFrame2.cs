using System;
using FinancialAnalytics.Wrappers.Office;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Converters;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interfaces;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class TextFrame2 : ExcelEntityWrapper<ITextFrame2>, ITextFrame2
    {
        #region Constants and Fields

		protected Microsoft.Office.Interop.Excel.TextFrame2 _textFrame2;
        private bool _disposed;

        #endregion

		public TextFrame2(ExcelEntityResolver entityResolver, Microsoft.Office.Interop.Excel.TextFrame2 textFrame2)
            : base(entityResolver)
        {
            if (textFrame2 == null)
            {
				throw new ArgumentNullException("textFrame2");
            }
			_textFrame2 = textFrame2;
        }

        public override bool Equals(ITextFrame2 obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
			var textFrame2 = (TextFrame2)obj;
			return _textFrame2.Equals(textFrame2._textFrame2);
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    //Here we must dispose managed resources
                }
                //Here we must dispose unmanaged resources and LOH objects
				ComObjectsFinalizer.ReleaseComObject(_textFrame2);
                _disposed = true;
            }
            base.Dispose(disposing);
        }

		public ITextRange2 TextRange
		{
			get { return EntityResolver.ResolveTextRange2(_textFrame2.TextRange); }
		}
	}
}