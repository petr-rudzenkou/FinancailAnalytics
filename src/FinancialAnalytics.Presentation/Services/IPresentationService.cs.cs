using System;

namespace FinancialAnalytics.Presentation.Services
{
    public interface IPresentationService
    {
        void Invoke(Action action);

        void BeginInvoke(Action action);

        void InvokeShutdown();

        void Dispose();
    }
}
