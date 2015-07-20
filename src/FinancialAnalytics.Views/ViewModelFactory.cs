using FinancialAnalytics.Core.Composition.Unity;
using FinancialAnalytics.Views.Charts.Interfaces;
using FinancialAnalytics.Views.HistoricalData.Interfaces;
using FinancialAnalytics.Views.LeagueTable.Interfaces;
using FinancialAnalytics.Views.Login.Interfaces;
using FinancialAnalytics.Views.Options.Interfaces;
using FinancialAnalytics.Views.Portfolio.Interfaces;
using FinancialAnalytics.Views.Quotes.Interfaces;
using FinancialAnalytics.Views.Screener.Interfaces;
using FinancialAnalytics.Views.Search.Interfaces;
using FinancialAnalytics.Views.XChangeRates.Interfaces;

namespace FinancialAnalytics.Views
{
    public class ViewModelFactory : IViewModelFactory
    {
        private readonly IServiceContainer _container;
        public ViewModelFactory(IServiceContainer container)
        {
            _container = container;
        }
        public IViewModel Create(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.StockScreener:
                    return _container.GetInstance<IScreenerViewModel>();
                case ViewType.LeagueTable:
                    return _container.GetInstance<ILeagueTableViewModel>();
                case ViewType.Portfolio:
                    return _container.GetInstance<IPortfolioViewModel>();
                case ViewType.Quotes:
                    return _container.GetInstance<IQuotesViewModel>();
                case ViewType.HistoricalData:
                    return _container.GetInstance<IHistoricalDataViewModel>();
                case ViewType.Charts:
                    return _container.GetInstance<IChartsViewModel>();
                case ViewType.Options:
                    return _container.GetInstance<IOptionsViewModel>();
                case ViewType.Search:
                    return _container.GetInstance<ISearchViewModel>();
                case ViewType.Login:
                    return _container.GetInstance<ILoginViewModel>();
                case ViewType.XChangeRates:
                    return _container.GetInstance<IXChangeRatesViewModel>();
                default:
                    return null;
            }
        }
    }
}
