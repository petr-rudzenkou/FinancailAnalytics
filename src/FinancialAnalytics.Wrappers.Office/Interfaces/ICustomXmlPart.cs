namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICustomXmlPart : IEntityWrapper<ICustomXmlPart>
	{
		/// <summary>
		/// Gets a String containing the GUID assigned to the current CustomXmlPart object.
		/// </summary>
		string Id { get; }

		/// <summary>
		/// Gets the XML representation of the current CustomXmlPart object.
		/// </summary>
		string Xml { get; }

		ICustomXmlPrefixMappings NamespaceManager { get; }

		ICustomXmlNodes SelectNodes(string xpath);

		ICustomXmlNode SelectSingleNode(string xpath);
	}
}
