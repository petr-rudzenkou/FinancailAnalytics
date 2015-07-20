using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IFont
    {
        object Size { get; set; }

        string Name { get; set; }

        Color Color { get; set; }

        object ColorIndex { get; set; }

        object ThemeColor { get; set; }
        object TintAndShade { get; set; }

        Color ChartColor { get; set; }

        /// <summary>
        /// True if the font is bold.
        /// </summary>
        bool Bold { get; set; }

        UnderlineStyle Underline { get; set; }

        object UnderlineAsObject { get; set; }

        bool Equals(IFont obj);

        bool Italic { get; set; }

        bool Strikethrough { get; set; }

        bool Subscript { get; set; }

        bool Superscript { get; set; }
    }
}
