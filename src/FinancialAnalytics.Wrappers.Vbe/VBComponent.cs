using System;
using FinancialAnalytics.Wrappers.Vbe.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe
{	
	internal class VBComponent : VbeEntityWrapper<IVBComponent>, IVBComponent
	{
		private Microsoft.Vbe.Interop.VBComponent _vbComponent;

		public VBComponent(VbeEntityResolver vbeEntityResolver, Microsoft.Vbe.Interop.VBComponent vbComponent)
			: base(vbeEntityResolver)
		{
			if (vbComponent == null)
			{
				throw new ArgumentNullException("vbComponent");
			}
			_vbComponent = vbComponent;
		}

		public override bool Equals(IVBComponent obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			VBComponent vbComponent = (VBComponent)obj;
			return this._vbComponent.Equals(vbComponent._vbComponent);
		}

		public object VBComponentObject
		{
			get { return _vbComponent; }
		}

		public ICodeModule CodeModule
		{
			get { return EntityResolver.ResolveCodeModule(_vbComponent.CodeModule); }
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
				ComObjectsFinalizer.ReleaseComObject(_vbComponent);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
		
	}
}
