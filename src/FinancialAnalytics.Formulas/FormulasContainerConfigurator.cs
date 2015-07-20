using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Core.Export;
using FinancialAnalytics.DataFacades.Quotes;


namespace FinancialAnalytics.Formulas
{
    public class FormulasContainerConfigurator : IConfigureContainer
    {
        public void ConfigureContainer(IServiceContainer container)
        {
            container.RegisterType<IFormulaHandler, FormulaHandler>(Lifetime.Singleton);
            container.RegisterType<IDataExporterFactory, DataExporterFactory>(Lifetime.Singleton);
        }
    }
}
