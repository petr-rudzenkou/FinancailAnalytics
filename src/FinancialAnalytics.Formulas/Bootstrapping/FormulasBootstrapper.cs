using System.Windows.Threading;
using DryTools.Primitives;
using FinancialAnalytics.Core;
using FinancialAnalytics.Core.Composition.Unity;
using Application = FinancialAnalytics.Wrappers.Excel.Application;

namespace FinancialAnalytics.Formulas.Bootstrapping
{
    public class FormulasBootstrapper
    {
        public FormulasBootstrapper(IServiceContainer container)
        {
            Container = container;
        }

        public IFormulaHandler FormulaHandler
        {
            get { return Container.GetInstance<IFormulaHandler>(); }
        }

        public IServiceContainer Container
        {
            get;
            private set;
        }

        public IApplicationProvider ApplicationProvider
        {
            get
            {
                return Container.GetInstance<IApplicationProvider>();
            }
        }

        public void Run(object application)
        {
            SetupRun();
            var applicationProvider = new ApplicationProvider(null);
            var app = new Application(application as Microsoft.Office.Interop.Excel.Application);
            applicationProvider.SetApplication(app);
            Container.RegisterType<IApplicationProvider>(Lifetime.Singleton,
                x => applicationProvider);

            Container.Configure(new FormulasContainerConfigurator());

        }

        private static void SetupRun()
        {
            var dispatcher = Dispatcher.CurrentDispatcher;
            DryTools.Execution.Run.Setup.IsOnUIChecker = dispatcher.CheckAccess;

            DryTools.Execution.Run.Setup.GetOnUIRunner = () => action =>
            {
                dispatcher.BeginInvoke(action);
                return Disposable.None;
            };

            DryTools.Execution.Run.Setup.GetOnUIAndWaitRunner = () => action =>
            {
                if (dispatcher.CheckAccess())
                    action();
                else
                    dispatcher.Invoke(action);
            };
        }

    }
}
