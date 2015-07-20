using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Presentation.UI.Window;

namespace FinancialAnalytics.Presentation.Core
{
    public  interface INavigatorFactory
    {
        INavigator GetNavigator();
    }
}
