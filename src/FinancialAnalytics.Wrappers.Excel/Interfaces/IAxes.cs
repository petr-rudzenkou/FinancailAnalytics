using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IAxes : IEntitiesCollectionWrapper<IAxes, IAxis>
    {
        IAxis Item(AxisType type);
    }
}
