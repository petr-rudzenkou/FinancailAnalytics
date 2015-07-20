using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;
using FinancialAnalytics.Wrappers.Office.Interfaces;

namespace FinancialAnalytics.Wrappers.Excel.Interfaces
{
    public interface IValidation : IEntityWrapper<IValidation>
    {
        string ErrorMessage { get; set; }
        string InputMessage { get; set; }
        string InputTitle { get; set; }
        bool IgnoreBlank { get; set; }
        void Add(XlDVType Type, object AlertStyle, object Operator, object Formula1, object Formula2);
        void Delete();
    }
}
