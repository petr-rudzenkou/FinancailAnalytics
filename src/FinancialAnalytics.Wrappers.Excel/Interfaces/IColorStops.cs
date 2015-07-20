using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IColorStops : IEntitiesCollectionWrapper<IColorStops, IColorStop>
	{
		void Clear();
		IColorStop Add(double position);
	}
}