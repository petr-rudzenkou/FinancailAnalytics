using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Converters;
using FinancialAnalytics.Wrappers.Office.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class Font2 : EntityWrapperBase<IFont2>, IFont2
	{
		protected Microsoft.Office.Core.Font2 _officeFont2;

		public Font2(EntityResolverBase entityResolver, Microsoft.Office.Core.Font2 font2)
			: base(entityResolver)
		{
			if (font2 == null)
			{
				throw new ArgumentNullException("font2");
			}
			_officeFont2 = font2;
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
				ComObjectsFinalizer.ReleaseComObject(_officeFont2);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(IFont2 obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			Font2 font2 = (Font2)obj;
			return _officeFont2.Equals(font2._officeFont2);
		}

		public TextUnderlineType UnderlineStyle
		{
			get
			{
				return MsoTextUnderlineTypeToTextUnderlineTypeConverter.Convert(_officeFont2.UnderlineStyle);				
			}
			set
			{
				_officeFont2.UnderlineStyle = MsoTextUnderlineTypeToTextUnderlineTypeConverter.ConvertBack(value);
			}
		}
	}
}
