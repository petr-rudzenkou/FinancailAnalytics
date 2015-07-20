using System;
using System.Net;

namespace FinancialAnalytics.DataFacades.Login
{
    public class YAccountManager
    {
        private bool _isLoggedIn = false;
        private CookieContainer _cookieContainer;

        public CookieContainer Cookies
        {
            get { return _cookieContainer; }
            set { _cookieContainer = value; }
        }

        public bool IsLoggedIn
        {
            get { return _isLoggedIn; }
            set { _isLoggedIn = value; }
        }
        public bool LogIn(NetworkCredential user)
        {
            if (!this.IsLoggedIn)
            {
                if (user == null) throw new ArgumentNullException("User credential is null.");
                Cookies = new CookieContainer();
                //TODO: make a request
                if (!this.IsLoggedIn) Cookies = null;
            }
            return this.IsLoggedIn;
        }

        public bool LogOut()
        {
            if (this.IsLoggedIn)
            {
                //Html2XmlDownload dl = new Html2XmlDownload();
                //dl.Settings.Account = this;
                //dl.Settings.DownloadStream = false;
                //dl.Settings.Url = "http://login.yahoo.com/config/login?logout=1&.direct=2&.done=&.src=&.intl=us&.lang=en-US";
                //dl.AsyncDownloadCompleted += this.LogOutAsync_Completed;
                //Response<XDocument> resp = dl.Download();
                //if (resp.Connection.State == ConnectionState.Success)
                //{
                //    mCookies = null;
                //    this.SetCrumb(string.Empty);
                //    if (this.PropertyChanged != null) this.PropertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs("IsLoggedIn"));
                //}
            }
            return this.IsLoggedIn;
        }
    }
}
