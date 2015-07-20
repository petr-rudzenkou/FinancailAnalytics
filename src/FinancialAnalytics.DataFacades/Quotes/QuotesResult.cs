namespace FinancialAnalytics.DataFacades.Quotes
{
    public class QuotesResult
    {
        private QuotesData[] mItems = null;
        public QuotesData[] Items
        {
            get { return mItems; }
        }
        internal QuotesResult(QuotesData[] items)
        {
            mItems = items;
        }
    }
}
