using System;

namespace FinancialAnalytics.Core.Formulas
{
    public interface IRefreshFormulasTimer
    {
        int UpdateInterval { get; set; }
        event EventHandler AutoUpdate;
        void Start();
        void Stop();

        bool Enabled { get; set; }

        void ResetInterval();
    }
}
