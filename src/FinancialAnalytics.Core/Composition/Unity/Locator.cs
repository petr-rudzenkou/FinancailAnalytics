using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialAnalytics.Core.Composition.Unity
{
    public static class Locator
    {
        static Locator()
        {
            Current = ServiceLocator.Create();
            Current.Container.RegisterInstance<Microsoft.Practices.ServiceLocation.IServiceLocator>(Current);
        }

        public static IServiceLocator Current { get; private set; }
    }
}
