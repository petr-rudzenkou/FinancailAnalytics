using FinancialAnalytics.DataFacades.Login;

namespace FinancialAnalytics.DataFacades.Base
{
    public class Html2XmlDownloadSettings : SettingsBase
    {
        public YAccountManager Account { get; set; }
        public string Url { get; set; }
        public bool DownloadStream { get; set; }
        protected override System.Net.CookieContainer Cookies { get { return Account != null ? this.Account.Cookies : null; } }
        protected override bool DownloadResponseStream { get { return DownloadStream; } }

        public Html2XmlDownloadSettings()
        {
            DownloadStream = true;
        }

        public override string GetUrl()
        {
            return Url;
        }

        public override object Clone()
        {
            return new Html2XmlDownloadSettings() { Account = this.Account, Url = this.Url };
        }
    }
}
