using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class TextRange2 : EntityWrapperBase<ITextRange2>, ITextRange2
	{
		protected Microsoft.Office.Core.TextRange2 _officeTextRange2;

		public TextRange2(EntityResolverBase entityResolver, Microsoft.Office.Core.TextRange2 textRange2)
			: base(entityResolver)
		{
			if (textRange2 == null)
			{
				throw new ArgumentNullException("textRange2");
			}
			_officeTextRange2 = textRange2;
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
				ComObjectsFinalizer.ReleaseComObject(_officeTextRange2);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(ITextRange2 obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			TextRange2 textRange2 = (TextRange2)obj;
			return _officeTextRange2.Equals(textRange2._officeTextRange2);
		}

		public IFont2 Font
		{
			get { return EntityResolver.ResolveFont2(_officeTextRange2.Font); }
		}
	}
}
