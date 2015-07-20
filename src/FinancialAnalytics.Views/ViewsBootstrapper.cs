using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using Caliburn.Micro;
using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Core.Export;
using FinancialAnalytics.DataFacades.Quotes;
using FinancialAnalytics.Views.Base;
using FinancialAnalytics.Views.Charts;
using FinancialAnalytics.Views.Charts.Interfaces;
using FinancialAnalytics.Views.ExcelExport;
using FinancialAnalytics.Views.HistoricalData;
using FinancialAnalytics.Views.HistoricalData.Interfaces;
using FinancialAnalytics.Views.LeagueTable;
using FinancialAnalytics.Views.LeagueTable.Interfaces;
using FinancialAnalytics.Views.Login;
using FinancialAnalytics.Views.Login.Interfaces;
using FinancialAnalytics.Views.Options;
using FinancialAnalytics.Views.Options.Interfaces;
using FinancialAnalytics.Views.Portfolio;
using FinancialAnalytics.Views.Portfolio.Base;
using FinancialAnalytics.Views.Portfolio.Interfaces;
using FinancialAnalytics.Views.ProgressBar;
using FinancialAnalytics.Views.Quotes;
using FinancialAnalytics.Views.Quotes.Base;
using FinancialAnalytics.Views.Quotes.Interfaces;
using FinancialAnalytics.Views.Screener;
using FinancialAnalytics.Views.Screener.Base;
using FinancialAnalytics.Views.Screener.Interfaces;
using FinancialAnalytics.Views.Search;
using FinancialAnalytics.Views.Search.Interfaces;
using FinancialAnalytics.Views.ViewSettings;
using FinancialAnalytics.Views.XChangeRates;
using FinancialAnalytics.Views.XChangeRates.Interfaces;

namespace FinancialAnalytics.Views
{
    public class ViewsBootstrapper : BootstrapperBase
    {
        public IServiceContainer Container
        {
            get;
            private set;
        }

        public ViewsBootstrapper(IServiceContainer container)
        {
            Container = container;
        }

        public void Run()
        {
            ConfigureContainer();
            RunWpfApplication();
            Start();
        }

        private void ConfigureContainer()
        {
            Container.RegisterType<IViewsRenderer, ViewsRenderer>(Lifetime.Singleton);
            Container.RegisterType<IWindowSettingsFactory, WindowSettingsFactory>(Lifetime.Singleton);
            Container.RegisterType<IViewModelFactory, ViewModelFactory>(Lifetime.Singleton);

            //Common
            Container.RegisterType<IProgressBarService, ProgressBarService>(Lifetime.Singleton);
            Container.RegisterType<IDataExporterFactory, DataExporterFactory>(Lifetime.Singleton);

            Container.RegisterType<ProgressBarView, ProgressBarView>();
            Container.RegisterType<IProgressBarViewModel, ProgressBarViewModel>();

            //Screener
            Container.RegisterType<ScreenerView, ScreenerView>(Lifetime.Singleton);
            Container.RegisterType<IScreenerViewModel, ScreenerViewModel>(Lifetime.Singleton);

            Container.RegisterType<ScreenBuilderView, ScreenBuilderView>();
            Container.RegisterType<IScreenBuilderViewModel, ScreenBuilderViewModel>();

            Container.RegisterType<ScreenResultsView, ScreenResultsView>();
            Container.RegisterType<IScreenResultsViewModel, ScreenResultsViewModel>();
            Container.RegisterType<IScreenerResultsCollection, ScreenerResultsCollection>(Lifetime.Singleton);

            //Portfolio
            Container.RegisterType<PortfolioView, PortfolioView>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioViewModel, PortfolioViewModel>(Lifetime.Singleton);
            Container.RegisterType<PortfolioBasicView, PortfolioBasicView>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioBasicViewModel, PortfolioBasicViewModel>(Lifetime.Singleton);
            Container.RegisterType<PortfolioDetailedView, PortfolioDetailedView>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioDetailedViewModel, PortfolioDetailedViewModel>(Lifetime.Singleton);
            Container.RegisterType<PortfolioFundamentalsView, PortfolioFundamentalsView>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioFundamentalsViewModel, PortfolioFundamentalsViewModel>(Lifetime.Singleton);
            Container.RegisterType<PortfolioPerformanceView, PortfolioPerformanceView>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioPerformanceViewModel, PortfolioPerformanceViewModel>(Lifetime.Singleton);
            Container.RegisterType<IPortfolioQuotesCollection, PortfolioQuotesCollection>(Lifetime.Singleton);
            
            //Historical data
            Container.RegisterType<HistoricalDataView, HistoricalDataView>(Lifetime.Singleton);
            Container.RegisterType<IHistoricalDataViewModel, HistoricalDataViewModel>(Lifetime.Singleton);

            //League Table
            Container.RegisterType<LeagueTableView, LeagueTableView>(Lifetime.Singleton);
            Container.RegisterType<ILeagueTableViewModel, LeagueTableViewModel>(Lifetime.Singleton);
            
            //Charts
            Container.RegisterType<ChartsView, ChartsView>(Lifetime.Singleton);
            Container.RegisterType<IChartsViewModel, ChartsViewModel>(Lifetime.Singleton);

            //Quotes
            Container.RegisterType<QuotesView, QuotesView>(Lifetime.Singleton);
            Container.RegisterType<IQuotesViewModel, QuotesViewModel>(Lifetime.Singleton);
            Container.RegisterType<IQuotesCollection, QuotesCollection>(Lifetime.Singleton);
            Container.RegisterType<QuotesInfoView, QuotesInfoView>();

            //Options
            Container.RegisterType<OptionsView, OptionsView>(Lifetime.Singleton);
            Container.RegisterType<IOptionsViewModel, OptionsViewModel>(Lifetime.Singleton);

            //Search
            Container.RegisterType<SearchView, SearchView>(Lifetime.Singleton);
            Container.RegisterType<ISearchViewModel, SearchViewModel>(Lifetime.Singleton);

            //Login
            Container.RegisterType<LoginView, LoginView>(Lifetime.Singleton);
            Container.RegisterType<ILoginViewModel, LoginViewModel>(Lifetime.Singleton);

            //XChange Rates
            Container.RegisterType<XChangeRatesView, XChangeRatesView>(Lifetime.Singleton);
            Container.RegisterType<IXChangeRatesViewModel, XChangeRatesViewModel>(Lifetime.Singleton);


            //Excel exporter
            Container.RegisterType<IExcelExporter, ExcelExporter>(Lifetime.Singleton);
        }

        private void RunWpfApplication()
        {
            //Create application object manually as wpf is hosted
            if (Application.Current == null)
            {
                var application = new Application();
                application.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }
        }

        protected override IEnumerable<Assembly> SelectAssemblies()
        {
            return new[] { Assembly.GetExecutingAssembly() };
        }

        protected override object GetInstance(Type serviceType, string key)
        {
            try
            {
                if (string.IsNullOrEmpty(key))
                    return Container.GetInstance(serviceType);
                return Container.GetInstance(serviceType, key);

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not locate any instances of contract {0}.", key ?? serviceType.Name));
            }
        }

        protected override IEnumerable<object> GetAllInstances(Type serviceType)
        {
            return Container.GetAllInstances(serviceType);
        }
    }
}
