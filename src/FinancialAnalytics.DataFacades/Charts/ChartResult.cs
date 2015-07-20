namespace FinancialAnalytics.DataFacades.Charts
{
    public class ChartResult
    {
        private string _id = null;
        private System.IO.MemoryStream mItem = null;

        public string Id
        {
            get { return _id; }
        }
        public System.IO.MemoryStream Item
        {
            get { return mItem; }
        }

        internal ChartResult()
        { 
        }
        internal ChartResult(string id, System.IO.MemoryStream item)
        {
            _id = id;
            mItem = item;
        }
    }
}
