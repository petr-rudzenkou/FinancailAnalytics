using Caliburn.Micro;
using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Presentation.Core;
using FinancialAnalytics.Presentation.Services;

namespace FinancialAnalytics.Presentation
{
    public class PresentationContainerConfigurator : IConfigureContainer
    {
        private IServiceContainer _container;
        public void ConfigureContainer(IServiceContainer container)
        {
            _container = container;
            _container.RegisterType<IOfficeWindowManager, OfficeWindowManager>(Lifetime.Singleton);
            _container.RegisterType<IEventAggregator, EventAggregator>(Lifetime.Singleton);
            _container.RegisterType<INavigatorFactory, NavigatorFactory>(Lifetime.Singleton);
            _container.RegisterType<INavigationService, NavigationService>(Lifetime.Singleton);
            _container.RegisterType<IPresentationService, PresentationService>(Lifetime.Singleton);
        }
    }
}
