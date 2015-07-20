using System.Collections.Generic;

namespace FinancialAnalytics.DataFacades.Login
{
    public class LoginDownloadSettings
    {
        public YAccountManager Account { get; set; }
        public string Url { get; set; }
        public string RefererUrlPart { get; set; }
        public List<KeyValuePair<string, string>> AdditionalWebForms { get; set; }
        public string[] SearchForWebForms { get; set; }
        public string FormActionPattern { get; set; }
        public bool DownloadResponse { get; set; }
    }
}
