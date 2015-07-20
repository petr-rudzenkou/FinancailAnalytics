namespace FinancialAnalytics.DataFacades.HistoricalData
{
    public class HistoricalDataResult
    {
        private HistoricalData[] mItems = null;
        public HistoricalData[] Items
        {
            get { return mItems; }
        }
        internal HistoricalDataResult(HistoricalData[] items)
        {
            mItems = items;
        }
    }
}
