using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinancialAnalytics.Core.Formulas
{
    public interface IDailyRefreshTimer
    {
        int? Sec { get; set; }
        int? Min { get; set; }
        int? Hour { get; set; }

        event EventHandler AutoUpdate;
        void Start();
        void Stop();

        bool Enabled { get; set; }
    }
}
