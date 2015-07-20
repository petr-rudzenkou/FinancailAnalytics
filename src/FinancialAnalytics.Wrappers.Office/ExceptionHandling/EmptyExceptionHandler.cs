using System;
using System.Diagnostics;

namespace FinancialAnalytics.Wrappers.Office.ExceptionHandling
{
    public class EmptyExceptionHandler : IExceptionHandler
    {
        public bool HandleException(Exception exception)
        {
            return false;
        }

        public void LogException(Exception exception)
        {
            
        }
    }
}
