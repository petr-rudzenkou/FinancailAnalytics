using System.Collections.Generic;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICustomXmlParts : IEntityWrapper<ICustomXmlParts>, IEnumerable<ICustomXmlPart>
	{
		int Count { get; }

		ICustomXmlPart this[object id] { get; }

		ICustomXmlPart Add(string xml);

		ICustomXmlParts SelectByNamespace(string @namespace);

		ICustomXmlPart SelectById(string id);
	}
}
