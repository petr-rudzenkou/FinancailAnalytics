using System.Windows;
using FinancialAnalytics.Presentation.Core;

namespace FinancialAnalytics.Presentation.Services
{
    public class NavigationService : INavigationService
    {
        private readonly INavigatorFactory _navigatorFactory;
        public NavigationService(INavigatorFactory navigatorFactory)
        {
            _navigatorFactory = navigatorFactory;
            //Cef.Initialize(new CefSettings());
        }
        public void Navigate(string url)
        {
            var navigator = _navigatorFactory.GetNavigator();
            ((Window) navigator).MakeProcessOwned();
            navigator.Navigate(url);
        }
    }
}
