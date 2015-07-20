using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IVisualBasicEditorWindow
    {
        WindowState WindowState { get; set; }

        bool Visible { get; set; }

        bool Equals(IVisualBasicEditorWindow obj);
    }
}
