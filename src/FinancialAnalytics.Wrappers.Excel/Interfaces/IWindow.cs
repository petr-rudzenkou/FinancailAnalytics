using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWindow : IEntityWrapper<IWindow>
    {
        double Zoom { get; set; }
        bool DisplayGridlines { get; set; }
        void Activate();
        bool Visible { get; set; }
        object WindowObject { get; }
        string Caption { get; }
        ISheets Sheets { get; }
        IRange VisibleRange { get; }
        IPanes Panes { get; }
		ISheet ActiveSheet { get; }

        int PointsToScreenPixelsX(int points);
        int PointsToScreenPixelsY(int points);
		void ScrollIntoView(int left, int top, int width, int height);
        int ScrollRow { get; set; }
        int ScrollColumn { get; set; }
        int SplitRow { get; set; }
        int SplitColumn { get; set; }
        bool FreezePanes { get; set; }
    }
}
