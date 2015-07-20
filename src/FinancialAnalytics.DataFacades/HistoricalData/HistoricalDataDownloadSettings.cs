using System;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.HistoricalData
{
    public class HistoricalDataDownloadSettings : SettingsBase
    {
        public Encoding TextEncoding { get; set; }

        public string[] IDs { get; set; }

        public HistoricalDataDownloadSettings()
        {
            this.TextEncoding = Encoding.UTF8;
        }

        public DateTime StartDate { get; set; } //TODO:Use DateTime
        public DateTime EndDate { get; set; }

        public override string GetUrl()
        {
            StringBuilder query = new StringBuilder();
            query.Append("select * from yahoo.finance.historicaldata where symbol in (");
            for (int i = 0; i < IDs.Length; i++)
            {
                query.Append('"');
                query.Append(IDs[i]);
                query.Append('"');
                if (i != IDs.Length - 1)
                {
                    query.Append(",");
                }
            }

            var startDateUsFormat = StartDate.ToString("yyyy-MM-dd");
            var endDateUsFormat = EndDate.ToString("yyyy-MM-dd");
            query.Append(")");
            query.Append(" and startDate = ");
            query.Append('"' + startDateUsFormat + '"');
            query.Append(" and endDate = ");
            query.Append('"' + endDateUsFormat + '"');

            string url = DataFacadesHelper.YqlUrl(query.ToString());
            return url;
        }

        public override object Clone()
        {
            HistoricalDataDownloadSettings cln = new HistoricalDataDownloadSettings();
            cln.IDs = (string[])this.IDs.Clone();
            return cln;
        }
    }
}
