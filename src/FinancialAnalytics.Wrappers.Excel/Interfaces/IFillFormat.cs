using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IFillFormat
    {
        bool Visible { get; set; }
        IColorFormat ForeColor { get; }
		float Transparency { get; set; }
    }
}
