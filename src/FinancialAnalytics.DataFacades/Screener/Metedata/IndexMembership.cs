namespace FinancialAnalytics.DataFacades.Screener.Metedata
{
    public class IndexMembership
    {
        private readonly string _id;
        private readonly string _name;
        private readonly string _displayName;

        public IndexMembership(string id, string name, string displayName)
        {
            _id = id;
            _name = name;
            _displayName = displayName;
        }
        public string Id { get { return _id; } }
        public string Name { get { return _name; } }
        public string DisplayName { get { return _displayName; } }
    }
}
