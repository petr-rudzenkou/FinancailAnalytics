using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IPlotArea
    {
        double Width { get; set; }
        double Height { get; set; }
        double Left { get; set; }
        double Top { get; set; }
        IChart Chart { get; }
        IBorder Border { get; }
        IInterior Interior { get; }
        IChartFillFormat Fill { get; }
        object Select();
		IChartFormat Format { get; }
    }
}
