using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IDataLabel
    {
        double Left { get; set; }
        double Top { get; set; }
        object Separator { get; set; }
        object Type { get; set; }
        bool ShowBubbleSize { get; set; }
        bool ShowCategoryName { get; set; }
        bool ShowLegendKey { get; set; }
        bool ShowPercentage { get; set; }
        bool ShowSeriesName { get; set; }
        bool ShowValue { get; set; }
        IBorder Border { get; }
        IInterior Interior { get; }
        IChartFillFormat Fill { get; }
        string Text { get; set; }
        object Select();
        Wrappers.Excel.Interfaces.IFont Font { get; }
        DataLabelPosition Position { get; set; }
		IChartFormat Format { get; }
    }
}
