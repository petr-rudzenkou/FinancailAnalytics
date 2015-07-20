namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICustomXmlPrefixMappings : IEntityWrapper<ICustomXmlPrefixMappings>
	{
		void AddNamespace(string prefix, string uri);
	}
}
