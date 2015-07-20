using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IChartObjects : IEntitiesCollectionWrapper<IChartObjects, IChartObject>
    {
        IChartObject Add(double left, double top, double width, double height);
    }
}
