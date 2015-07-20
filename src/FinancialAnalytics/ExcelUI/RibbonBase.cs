using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;

namespace FinancialAnalytics.ExcelUI
{
    public class RibbonBase
    {
        private readonly List<IRibbonElement> _ribbonElements = new List<IRibbonElement>();

        protected List<IRibbonElement> RibbonElements
        {
            get { return _ribbonElements; }
        }

        protected IRibbonElement FindRibbonElement(IRibbonControl control)
        {
            return RibbonElements.FirstOrDefault(x => x.Id == control.Id);
        }
    }
}
