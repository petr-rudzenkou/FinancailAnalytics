using FinancialAnalytics.Core.Composition.Unity;

namespace FinancialAnalytics.AuthenticationClient
{
    public class AuthenticationContainerConfigurator : IConfigureContainer
    {
        public void ConfigureContainer(IServiceContainer container)
        {
            container.RegisterType<IAuthenticationClient, AuthenticationClient>(Lifetime.Singleton);
        }
    }
}
