using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IInterior : IEntityWrapper<IInterior>
    {
        object ColorIndex { get; set; }
        object Pattern { get; set; }
        object Color { get; set; }
        object PatternColorIndex { get; set; }
        object PatternColor { get; set; }
        object ThemeColor { get; set; }
        object TintAndShade { get; set; }
        IGradient Gradient { get; }
        Pattern GetPattern();
    }
}
