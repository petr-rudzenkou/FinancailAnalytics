using System;
using FinancialAnalytics.Wrappers.Vbe.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe
{	
	internal class VBProject : VbeEntityWrapper<IVBProject>, IVBProject
	{
		private Microsoft.Vbe.Interop.VBProject _vbProject;

		public VBProject(VbeEntityResolver vbeEntityResolver, Microsoft.Vbe.Interop.VBProject vbProject)
			: base(vbeEntityResolver)
		{
			if (vbProject == null)
			{
				throw new ArgumentNullException("vbProject");
			}
			_vbProject = vbProject;
		}

		public override bool Equals(IVBProject obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			VBProject vbProject = (VBProject)obj;
			return this._vbProject.Equals(vbProject._vbProject);
		}

		public IVBComponents VBComponents
		{
			get { return EntityResolver.ResolveVBComponents(_vbProject.VBComponents); }
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
				ComObjectsFinalizer.ReleaseComObject(_vbProject);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
		
	}
}
