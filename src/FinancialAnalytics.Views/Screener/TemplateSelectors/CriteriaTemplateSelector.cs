using System.Windows;
using System.Windows.Controls;
using FinancialAnalytics.DataFacades.Screener.Criterias;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.Views.Screener.TemplateSelectors
{
    public class CriteriaTemplateSelector : DataTemplateSelector
    {
        public DataTemplate IndustryCriteriaTemplate { get; set; }
        public DataTemplate RangeCriteriaTemplate { get; set; }
        
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            var industryCriteria = item as IndustryCriteria;
            if (industryCriteria != null)
            {
                return IndustryCriteriaTemplate;
            }
            var rangeCriteria = item as RangeCriteria;
            if (rangeCriteria != null)
            {
                return RangeCriteriaTemplate;
            }
            return null;
        }
    }
}
