using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IChartTitle
    {
        string Text { get; set; }

        void Delete();

        bool Equals(IChartTitle obj);

        object Orientation { get; set; }

        double Left { get; set; }

        double Top { get; set; }

        IFont Font { get; }

		IChartFillFormat Fill { get; }

		IChartFormat Format { get; }

        IBorder Border { get; }
    }
}
