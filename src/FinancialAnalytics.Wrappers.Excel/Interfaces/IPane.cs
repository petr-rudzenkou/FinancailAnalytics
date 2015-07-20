using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPane : IEntityWrapper<IPane>
    {
        int Index { get; }
        object PaneObject { get; }
        IRange VisibleRange { get; }
        int ScrollColumn { get; }
        int ScrollRow { get; }
        int PointsToScreenPixelsX (int Points);
        int PointsToScreenPixelsY(int Points);
    }
}
