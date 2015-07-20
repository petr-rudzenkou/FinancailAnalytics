using System;

namespace FinancialAnalytics.Wrappers.Office.ExceptionHandling
{
    public interface IExceptionHandler
    {
        bool HandleException(Exception exception);

        void LogException(Exception exception);
    }
}
