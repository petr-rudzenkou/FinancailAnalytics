namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICustomXmlNode : IEntityWrapper<ICustomXmlNode>
	{
		string Xml { get; }

		void AppendChildSubtree(string xml);

		ICustomXmlNodes SelectNodes(string xpath);

		ICustomXmlNode SelectSingleNode(string xpath);

		void Delete();
	}
}
