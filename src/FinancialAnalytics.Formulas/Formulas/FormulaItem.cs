using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Excel.Interfaces;

namespace FinancialAnalytics.Formulas.Formulas
{
    public class FormulaItem
    {
        public string[] Symbols { get; set; }
        public string[] DataItems { get; set; }
        public string Layout { get; set; }
        public IRange DestinationCell { get; set; }
        public string Address { get; set; }
        public bool Handled { get; set; }
        public bool Error { get; set; }
        public bool WithDestinationCell { get; set; }
    }
}
