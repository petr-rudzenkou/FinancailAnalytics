using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Presentation.UI.Window;

namespace FinancialAnalytics.Presentation.Core
{
    public class NavigatorFactory : INavigatorFactory
    {
        public INavigator GetNavigator()
        {
            var navigator = new Navigator();
            return navigator;
        }
    }
}
