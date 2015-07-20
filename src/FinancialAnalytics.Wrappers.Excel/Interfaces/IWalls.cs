using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IWalls
    {
        IInterior Interior { get; }
        IChartFormat Format { get; }
        IBorder Border { get; }
    }
}
