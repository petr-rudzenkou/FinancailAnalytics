using System.Windows.Threading;
using DryTools.Primitives;
using FinancialAnalytics.AuthenticationClient;
using FinancialAnalytics.Core;
using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Core.Formulas;
using FinancialAnalytics.ExcelUI;
using FinancialAnalytics.Presentation;
using FinancialAnalytics.Presentation.Core;
using FinancialAnalytics.Presentation.Services;
using FinancialAnalytics.Utils;
using FinancialAnalytics.Utils.Options;
using FinancialAnalytics.Views;
using FinancialAnalytics.Wrappers.Excel;


namespace FinancialAnalytics.Bootstrapping
{
    public class CommonBootstrapper
    {
        private readonly ViewsBootstrapper _viewsBootstrapper;

        public CommonBootstrapper(IServiceContainer container)
        {
            Container = container;
            _viewsBootstrapper = Container.GetInstance<ViewsBootstrapper>();
        }

        public IRibbon Ribbon
        {
            get
            {
                return Container.GetInstance<IRibbon>();
            }
        }

        public IOfficeWindowManager WindowManager
        {
            get
            {
                return Container.GetInstance<IOfficeWindowManager>();
            }
        }

        public IPresentationService PresentationService
        {
            get
            {
                return Container.GetInstance<IPresentationService>();
            }
        }
        
        public IApplicationProvider ApplicationProvider
        {
            get
            {
                return Container.GetInstance<IApplicationProvider>();
            }
        }

        public IServiceContainer Container
        {
            get;
            private set;
        }

        public FinancialAnalytics.Core.Composition.Unity.ServiceContainer ServiceContainer
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public FinancialAnalytics.ExcelUI.Ribbon Ribbon1
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }


        public void Run(object application)
        {
            SetupRun();

            var applicationProvider = new ApplicationProvider(null);
            var app = new Application(application as Microsoft.Office.Interop.Excel.Application);
            applicationProvider.SetApplication(app);
            Container.RegisterType<IApplicationProvider>(Lifetime.Singleton, x => applicationProvider);

            Container.RegisterType<IRefreshFormulasTimer, RefreshFormulasTimer>(Lifetime.Singleton);
            Container.RegisterType<IDailyRefreshTimer, DailyRefreshTimer>(Lifetime.Singleton);

            Container.Configure(new ExcelUIContainerConfigurator());
            Container.Configure(new PresentationContainerConfigurator());
            Container.Configure(new AuthenticationContainerConfigurator());

            _viewsBootstrapper.Run();
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
