using System.Windows;
using System.Windows.Controls;
using FinancialAnalytics.Utils.Options;

namespace FinancialAnalytics.Views.Options.TemplateSelectors
{
    class OptionsDataTemplateSelector : DataTemplateSelector
    {
        public DataTemplate RefreshFrequencyOptionDataTemplate { get; set; }
        public DataTemplate DailyRefreshTimeOptionDataTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            var refreshFrequencyOption = item as RefreshFrequencyOption;
            if (refreshFrequencyOption != null)
            {
                return RefreshFrequencyOptionDataTemplate; 
            }
            var dailyRefreshTimeOption = item as DailyRefreshTimeOption;
            if (dailyRefreshTimeOption != null)
            {
                return DailyRefreshTimeOptionDataTemplate;
            }
            return null;
        }
    }
}
