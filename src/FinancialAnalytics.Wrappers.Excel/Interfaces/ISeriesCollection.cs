using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ISeriesCollection : IEntitiesCollectionWrapper<ISeriesCollection, ISeries>
    {
        ISeries Add(IRange source, RowCol rowCol);

        ISeries Add(IRange source, RowCol rowCol, bool seriesLabels, bool categoryLabels, bool replace);

		ISeries NewSeries();
	}
}
