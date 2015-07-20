using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Vbe.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe
{
	public class VbeEntityResolver : EntityResolverBase
	{
		public IVBProject ResolveVBProject(Microsoft.Vbe.Interop.VBProject vbeVBProject)
		{
			return new VBProject(this, vbeVBProject);
		}

		public IVBComponents ResolveVBComponents(Microsoft.Vbe.Interop.VBComponents vbeVBComponents)
		{
			return new VBComponents(this, vbeVBComponents);
		}

		public IVBComponent ResolveVBComponent(Microsoft.Vbe.Interop.VBComponent vbeVBComponent)
		{
			return new VBComponent(this, vbeVBComponent);
		}

		public ICodeModule ResolveCodeModule(Microsoft.Vbe.Interop.CodeModule vbeCodeModule)
		{
			return new CodeModule(this, vbeCodeModule);
		}
	}
}