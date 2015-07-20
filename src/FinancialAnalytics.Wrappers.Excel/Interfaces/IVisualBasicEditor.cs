using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IVisualBasicEditor
    {
        IVisualBasicEditorWindow MainWindow { get; }

        bool Equals(IVisualBasicEditor obj);
    }
}
