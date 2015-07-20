using FinancialAnalytics.DataFacades.XChangeRates.Metadata;

namespace FinancialAnalytics.DataFacades.XChangeRates
{
    public class XChangeRatesResult
    {
        private XChangeRate[] mItems = null;
        public XChangeRate[] Items
        {
            get { return mItems; }
        }
        internal XChangeRatesResult(XChangeRate[] items)
        {
            mItems = items;
        }
    }
}
