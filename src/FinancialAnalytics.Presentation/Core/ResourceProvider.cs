using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace FinancialAnalytics.Presentation.Core
{
    public class ResourceProvider
    {
        public ResourceDictionary GetResourceDictionary()
        {
            return new ResourceDictionary()
            {
                Source =
                    new Uri(@"pack://application:,,,/FinancialAnalytics.Presentation;component/Styles/FinancialAnalyticsStylesAndTemplates.xaml")
                    //new Uri(@"pack://application:,,,/Nova;component/Resources/Window.xaml")
            };
        }
    }
}
