using System;
using Microsoft.Practices.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
    ///<summary>
    /// Wraps underlying container into custom interfaces used in Common Office Framework
    /// </summary>
    /// <remarks>
    /// Helps to be flexible in usage cases and performances optimizations.
    /// </remarks>
    ///<seealso href="http://philipm.at/2011/0819/#tldr"/>
	public interface IServiceContainer: Microsoft.Practices.ServiceLocation.IServiceLocator
	{
        IServiceContainer Configure(params IConfigureContainer[] containerConfigurators);

        IServiceContainer RegisterInstance<TInterface>(TInterface instance);
        IServiceContainer RegisterInstance<TInterface>(string name, TInterface instance);

        IServiceContainer RegisterInstance<TInterface>(TInterface instance, LifetimeManager lifetimeManager);
        IServiceContainer RegisterType<TFrom, TTo>() where TTo : TFrom;

		IServiceContainer RegisterType<TFrom, TTo>(LifetimeManager lifetimeManager) where TTo : TFrom;

        IServiceContainer RegisterType<T>(LifetimeManager lifetimeManager);

        IServiceContainer RegisterType<T>(LifetimeManager lifetimeManager,Func<IServiceLocator,T> factory);
        IServiceContainer RegisterType<T>(Func<IServiceLocator,T> factory);

        IServiceContainer RegisterType<TFrom, TTo>(string name) where TTo : TFrom;
        IServiceContainer RegisterType<TFrom, TTo>(string name, LifetimeManager lifetimeManager) where TTo : TFrom;
		
        void Teardown(object instance);
        
	}
}