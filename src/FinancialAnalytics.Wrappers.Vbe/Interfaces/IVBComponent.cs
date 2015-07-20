using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Vbe.Interfaces
{
	public interface IVBComponent : IEntityWrapper<IVBComponent>
    {
		object VBComponentObject { get; }
		ICodeModule CodeModule { get; }
    }
}
