using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ICubeField : IEntityWrapper<ICubeField>
    {
        bool EnableMultiplePageItems { get; set; }
        string CurrentPageName { get; set; }
        IPivotFields PivotFields { get; }
        PivotFieldOrientation Orientation { get; set; }
    }
}
