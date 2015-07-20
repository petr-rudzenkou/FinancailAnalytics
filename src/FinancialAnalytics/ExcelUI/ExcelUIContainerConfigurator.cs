using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.ExcelUI.Ribbons;

namespace FinancialAnalytics.ExcelUI
{
    public class ExcelUIContainerConfigurator : IConfigureContainer
    {
        public void ConfigureContainer(IServiceContainer container)
        {
            container.RegisterType<IRibbon, Ribbon>(Lifetime.Singleton);
            container.RegisterType<IDataToolsRibbon, DataToolsRibbon>(Lifetime.Singleton);
            container.RegisterType<IRefreshManager, RefreshManager>(Lifetime.Singleton);
        }
    }
}
