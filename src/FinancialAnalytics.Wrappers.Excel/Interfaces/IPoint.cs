using FinancialAnalytics.Wrappers.Office.Interfaces;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPoint : IEntityWrapper<IPoint>
    {
        bool HasDataLabel { get; set; }
        IDataLabel DataLabel { get; }
        IBorder Border { get; }
        IInterior Interior { get; }
        IChartFillFormat Fill { get; }
        object Select();
        int MarkerBackgroundColor { get; set; }
		int MarkerForegroundColor { get; set; }
        object Parent { get; }
        ColorIndex MarkerBackgroundColorIndex { get; set; }
		ColorIndex MarkerForegroundColorIndex { get; set; }
		MarkerStyle MarkerStyle { get; set; }
		int MarkerSize { get; set; }
    }
}
