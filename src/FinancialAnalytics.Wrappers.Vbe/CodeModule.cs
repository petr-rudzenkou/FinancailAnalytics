using System;
using FinancialAnalytics.Wrappers.Vbe.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe
{	
	internal class CodeModule : VbeEntityWrapper<ICodeModule>, ICodeModule
	{
		private Microsoft.Vbe.Interop.CodeModule _vbCodeModule;

		public CodeModule(VbeEntityResolver vbeEntityResolver, Microsoft.Vbe.Interop.CodeModule vbCodeModule)
			: base(vbeEntityResolver)
		{
			if (vbCodeModule == null)
			{
				throw new ArgumentNullException("CodeModule");
			}
			_vbCodeModule = vbCodeModule;
		}

		public override bool Equals(ICodeModule obj)
		{
			if (obj == null || GetType() != obj.GetType())
			{
				return false;
			}
			CodeModule vbCodeModule = (CodeModule)obj;
			return this._vbCodeModule.Equals(vbCodeModule._vbCodeModule);
		}
		
		public void AddFromString(string codeString)
		{
			_vbCodeModule.AddFromString(codeString);
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
				ComObjectsFinalizer.ReleaseComObject(_vbCodeModule);
				disposed = true;
			}
			base.Dispose(disposing);
		}
		#endregion
		
	}
}
