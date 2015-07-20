using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{
	internal class ColorFormat : EntityWrapperBase<IColorFormat>, IColorFormat
	{
		protected Microsoft.Office.Core.ColorFormat _officeColorFormat;

		public ColorFormat(EntityResolverBase entityResolver, Microsoft.Office.Core.ColorFormat colorFormat)
			: base(entityResolver)
		{
			if (colorFormat == null)
			{
				throw new ArgumentNullException("colorFormat");
			}
			_officeColorFormat = colorFormat;
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
				ComObjectsFinalizer.ReleaseComObject(_officeColorFormat);
				disposed = true;
			}
			base.Dispose(disposing);
		}

		#endregion

		public override bool Equals(IColorFormat obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			ColorFormat colorFormat = (ColorFormat)obj;
			return _officeColorFormat.Equals(colorFormat._officeColorFormat);
		}

		public int RGB
		{
			get
			{
				return _officeColorFormat.RGB;
			}
			set
			{
				_officeColorFormat.RGB = value;
			}
		}
	}
}
