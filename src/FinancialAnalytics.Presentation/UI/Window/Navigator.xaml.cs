using System;

namespace FinancialAnalytics.Presentation.UI.Window
{
    /// <summary>
    /// Interaction logic for Navigator.xaml
    /// </summary>
    public partial class Navigator : INavigator
    {
        //private IWebBrowser _cefBrowser;
        private string _currentUrl;

        public Navigator()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            try
            {
                //Cef.Initialize(new CefSettings());
                //_cefBrowser = new ChromiumWebBrowser();
                //Content = _cefBrowser;
            }
            catch (Exception ex)
            { }
        }

        public void Navigate(string url)
        {
            try
            {
                //_currentUrl = url;
                //_cefBrowser.Load(_currentUrl);
                Show();
            }
            catch (Exception ex)
            { }
        }
    }
}
