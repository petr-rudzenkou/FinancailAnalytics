using System;
using Microsoft.Practices.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
    public interface IServiceLocator : Microsoft.Practices.ServiceLocation.IServiceLocator
	{
		IServiceContainer Container { get; }
		TService GetInstance<TService>(params ResolverOverride[] overrides);
	}
}