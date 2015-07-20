using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.DataFacades.Base;

namespace FinancialAnalytics.DataFacades.XChangeRates
{
    public class XChangeRatesDownloadSettings : SettingsBase
    {
        public Encoding TextEncoding { get; set; }
        public string[] IDs { get; set; }
        public override string GetUrl()
        {
            StringBuilder query = new StringBuilder();
            query.Append("select * from yahoo.finance.xchange where pair in (");
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
            XChangeRatesDownloadSettings cln = new XChangeRatesDownloadSettings();
            cln.IDs = (string[])this.IDs.Clone();
            return cln;
        }
    }
}
