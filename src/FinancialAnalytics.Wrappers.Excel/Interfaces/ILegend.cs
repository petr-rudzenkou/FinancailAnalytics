using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface ILegend : IEntityWrapper<ILegend>
    {
        double Height { get; set; }
        double Left { get; set; }
        LegendPosition Position { get; set; }
        double Top { get; set; }
        double Width { get; set; }
        bool IncludeInLayout { get; set;  }
        Wrappers.Excel.Interfaces.IFont Font { get; }
        IInterior Interior { get; }
		IChartFormat Format { get; }
        IBorder Border { get; }
		ILegendEntries LegendEntries(object index);
    }
}
