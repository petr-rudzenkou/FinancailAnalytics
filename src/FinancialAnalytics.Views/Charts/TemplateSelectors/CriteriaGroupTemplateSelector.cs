using System.Windows;
using System.Windows.Controls;
using FinancialAnalytics.Views.Charts.CriteriaGroups;

namespace FinancialAnalytics.Views.Charts.TemplateSelectors
{
    public class CriteriaGroupTemplateSelector : DataTemplateSelector
    {
        public DataTemplate BasicGroupTemplate { get; set; }
        public DataTemplate OtherGroupTemplate { get; set; }
        public DataTemplate CompareVsGroupTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            var rangeGroup = item as RangeGroup;
            if (rangeGroup != null)
            {
                return BasicGroupTemplate;
            }
            var typeGroup = item as TypeGroup;
            if (typeGroup != null)
            {
                return BasicGroupTemplate;
            }
            var scaleGroup = item as ScaleGroup;
            if (scaleGroup != null)
            {
                return BasicGroupTemplate;
            }
            var sizeGroup = item as SizeGroup;
            if (sizeGroup != null)
            {
                return BasicGroupTemplate;
            }
            var compareVsGroup = item as CompareVsGroup;
            if (compareVsGroup != null)
            {
                return CompareVsGroupTemplate;
            }
            return OtherGroupTemplate;
        }
    }
}
