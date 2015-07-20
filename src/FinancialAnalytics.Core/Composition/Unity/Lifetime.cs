using System;
using Microsoft.Practices.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
  
    public static class Lifetime
    {
        /// <summary>
        /// Shorthand  for <see cref="ContainerControlledLifetimeManager"/>
        /// </summary>
        public static LifetimeManager Singleton
        {
            get
            {
                // we cannot reuse lifetime manager cause Unity throws exception then
                return new ContainerControlledLifetimeManager();
            }
        }
    }
}
