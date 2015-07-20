using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.Quotes
{
    public class QuotesDownloadSettings : SettingsBase
    {
        public Encoding TextEncoding { get; set; }

        public string[] IDs { get; set; }
        public QuoteProperty[] Properties = new QuoteProperty[]
        {
                QuoteProperty.Symbol,
                QuoteProperty.Name,
                QuoteProperty.Open,
                QuoteProperty.PreviousClose,
                QuoteProperty.Change,
                QuoteProperty.LastTradePriceOnly
        };

        public QuotesDownloadSettings()
        {
            this.TextEncoding = Encoding.UTF8;
        }

        public override string GetUrl()
        {
            StringBuilder query = new StringBuilder();
            query.Append("select * from yahoo.finance.quotes where symbol in (");
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
            query.Append(")");
            return DataFacadesHelper.YqlUrl(query.ToString());
        }

        public override object Clone()
        {
            QuotesDownloadSettings cln = new QuotesDownloadSettings();
            cln.IDs = (string[])this.IDs.Clone();
            cln.Properties = (QuoteProperty[])this.Properties.Clone();
            return cln;
        }
    }
}
