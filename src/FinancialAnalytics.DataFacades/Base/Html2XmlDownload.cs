using System.Xml.Linq;

namespace FinancialAnalytics.DataFacades.Base
{
    public class Html2XmlDownload : DownloadClient<XDocument>
    {
        public Html2XmlDownloadSettings Settings { get { return (Html2XmlDownloadSettings)base.Settings; } set { base.SetSettings(value); } }

        public Html2XmlDownload()
        {
            Settings = new Html2XmlDownloadSettings();
        }

        protected override XDocument ConvertResult(ConnectionInfo connInfo, System.IO.Stream stream, SettingsBase settings)
        {
            return DataFacadesHelper.ParseXmlDocument(stream);
        }
    }
}
