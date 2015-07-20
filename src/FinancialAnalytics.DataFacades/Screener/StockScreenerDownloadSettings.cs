using System.Text;
using FinancialAnalytics.DataFacades.Base;
using FinancialAnalytics.DataFacades.Screener.Criterias.Base;

namespace FinancialAnalytics.DataFacades.Screener
{
    public class StockScreenerDownloadSettings : SettingsBase
    {
        public Encoding TextEncoding { get; set; }

        public string IndustryId { get; set; }
        public Criteria[] Criterias { get; set; }

        public StockScreenerDownloadSettings(Criteria[] criterias)
        {
            this.TextEncoding = Encoding.UTF8;
            this.Criterias = criterias;
        }
       
        public override string GetUrl()
        {
            StringBuilder query = new StringBuilder();
            query.Append("select * from yahoo.finance.quotes where symbol in (select company.symbol from yahoo.finance.industry where id in(");
            query.Append(IndustryId);
            query.Append("))");
            string url = DataFacadesHelper.YqlUrl(query.ToString());
            return url;
        }

        public override object Clone()
        {
            StockScreenerDownloadSettings cln = new StockScreenerDownloadSettings(this.Criterias);
            cln.IndustryId = (string)this.IndustryId.Clone();
            cln.Criterias = (Criteria[])this.Criterias.Clone();
            return cln;
        }
    }
}
