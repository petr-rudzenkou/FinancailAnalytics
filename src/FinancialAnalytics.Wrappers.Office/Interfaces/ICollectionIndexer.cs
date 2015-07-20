namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
	public interface ICollectionIndexer<out T>
	{
		T this[int index] { get; }
		int Count { get; }
	}
}