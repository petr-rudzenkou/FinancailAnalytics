using FinancialAnalytics.DataFacades.Quotes;

namespace FinancialAnalytics.DataFacades.Screener
{
    public class StockScreenerResult
    {
        private QuotesData[] mItems = null;
        public QuotesData[] Items
        {
            get { return mItems; }
        }

        public StockScreenerResult()
        {
        }
        public StockScreenerResult(QuotesData[] items)
        {
            mItems = items;
        }

        public bool FinalResponse { get; set; }
    }
}
