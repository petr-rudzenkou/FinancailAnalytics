using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{

    public interface IBorder : IEntityWrapper<IBorder>
    {
        object Color { get; set; }
        object ColorIndex { get; set; }
        object TintAndShade { get; set; }
        object ThemeColor { get; set; }
        LineStyle LineStyle { get; set; }
        object LineStyleObject { get; set; }
        BorderWeight Weight { get; set; }
    }
}
