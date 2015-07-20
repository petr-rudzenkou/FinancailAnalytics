using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface ITheme
    {
        IThemeColorScheme ThemeColorScheme { get; }
        IThemeFontScheme ThemeFontScheme { get; }
    }
}
