
namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICustomXmlNodes : IEntityWrapper<ICustomXmlNodes>
	{
		int Count { get; }

		ICustomXmlNode this[int index] { get; }
	}
}
