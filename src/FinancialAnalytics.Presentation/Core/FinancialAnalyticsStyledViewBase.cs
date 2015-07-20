using System.Windows;
using System.Windows.Controls;

namespace FinancialAnalytics.Presentation.Core
{
    //TODO: Add resource prodiver
    public class FinancialAnalyticsStyledViewBase : UserControl
    {
        private readonly ResourceDictionary _d;

        //public FinancialAnalyticsStyledViewBase() :base()
        //{
            
        //    //var resourceProvider = new ResourceProvider();
        //    //_d = resourceProvider.GetResourceDictionary();
        //    CommonOfficeStyledViewBase_Loaded(null, null);
        //}

		protected override void OnInitialized(System.EventArgs e)
		{
			base.OnInitialized(e);
			Unloaded += CommonOfficeStyledViewBase_Unloaded;
			Loaded += CommonOfficeStyledViewBase_Loaded;
		}

        public static readonly DependencyProperty HeaderContentProperty = DependencyProperty.Register("HeaderContent", typeof(object), typeof(FinancialAnalyticsStyledViewBase), new FrameworkPropertyMetadata(null));

        public object HeaderContent
        {
            get
            {
				return GetValue(HeaderContentProperty);
            }
            set
            {
                SetValue(HeaderContentProperty, value);
            }
        }

        private void CommonOfficeStyledViewBase_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Resources.MergedDictionaries.Contains(_d))
                Resources.MergedDictionaries.Add(_d);
        }


        private void CommonOfficeStyledViewBase_Unloaded(object sender, RoutedEventArgs e)
        {
            if (Resources.MergedDictionaries.Contains(_d))
                Resources.MergedDictionaries.Remove(_d);
        }

        //public override void OnApplyTemplate()
        //{
            
        //}
    }
}
