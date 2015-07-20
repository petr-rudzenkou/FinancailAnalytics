using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office.Interfaces
{
    [ComVisible(true)]
    public interface ICOMAddIn
    {
        object Object { get; }
        bool Connect { get; }
    }
}
