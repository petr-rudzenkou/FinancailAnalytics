using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.ExcelUI
{
    public interface IRibbonElement
    {
        string Id { get; }

        string Label { get; }

        bool IsEnabled { get; }

        bool IsVisible { get; }

        string ScreenTip { get; }

        string ScreenSuperTip { get; }

        string KeyTip { get; }

        Action Action { get; }

        bool IsPressed { get; }

        string Content { get; }

        Bitmap Image { get; }
        Bitmap ImagesMask { get; }
    }
}
