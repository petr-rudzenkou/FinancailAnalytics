using System;
using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Vbe.Enums;

namespace FinancialAnalytics.Wrappers.Vbe.Interfaces
{
	public interface IVBComponents : IEntitiesCollectionWrapper<IVBComponents, IVBComponent>
	{
		IVBComponent Add(VbComponentType componentType);
		void Remove(IVBComponent vbComponent);
	}
}
