using System;

using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface ILegendKey : IEntityWrapper<ILegendKey>
	{
		IInterior Interior { get; }
		IBorder Border { get; }
		IChartFillFormat Fill { get; }
		int MarkerBackgroundColor { get; set; }
		int MarkerForegroundColor { get; set; }
		ColorIndex MarkerBackgroundColorIndex { get; set; }
		ColorIndex MarkerForegroundColorIndex { get; set; }
		MarkerStyle MarkerStyle { get; set; }
	}
}
