using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.ExcelUI
{
    public interface IRefreshManager
    {
        void Refresh(string refreshMode);
    }
}
