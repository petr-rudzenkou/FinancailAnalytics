using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
	public interface IChartArea : IEntityWrapper<IChartArea>
    {
        double Width { get; set; }

        double Height { get; set; }

        double Left { get; set; }

        double Top { get; set; }
        
        bool Shadow { get; set; }

        IBorder Border { get; }
        
        IChart Chart { get; }

        IFont Font { get; }

        IInterior Interior { get; }

        IChartFillFormat Fill { get; }

        void Select();

        void Copy();

		IChartFormat Format { get; }
    }
}
