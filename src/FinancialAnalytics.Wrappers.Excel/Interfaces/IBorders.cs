using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IBorders : IEntitiesCollectionWrapper<IBorders, IBorder>
    {
        object Weight { get; set; }
		object Color { get; set; }
		object ColorIndex { get; set; }
        object LineStyle { get; set; }
        IBorder this[BordersIndex bordersIndex] { get; }
    }
}
