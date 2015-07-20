using System;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Office
{	
	internal class GradientStop : EntityWrapperBase<IGradientStop>, IGradientStop
	{
		private Microsoft.Office.Core.GradientStop _officeGradientStop;

		public GradientStop(EntityResolverBase entityResolver, Microsoft.Office.Core.GradientStop gradientStop)
			: base(entityResolver)
		{
			if (gradientStop == null)
			{
				throw new ArgumentNullException("gradientStop");
			}
			_officeGradientStop = gradientStop;
		}

		public override bool Equals(IGradientStop obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			GradientStop gradientStop = (GradientStop)obj;
			return this._officeGradientStop.Equals(gradientStop._officeGradientStop);
		}

		public IColorFormat Color
		{
			get
			{
				return EntityResolver.ResolveColorFormat(_officeGradientStop.Color);
			}
		}

		public float Position
		{
			get 
			{
				return _officeGradientStop.Position;
			}
			set 
			{
				_officeGradientStop.Position = value;
			}
		}

		public float Transparency
		{
			get
			{
				return _officeGradientStop.Transparency; 
			}
			set
			{
				_officeGradientStop.Transparency = value;
			}
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
				ComObjectsFinalizer.ReleaseComObject(_officeGradientStop);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
		
	}
}
