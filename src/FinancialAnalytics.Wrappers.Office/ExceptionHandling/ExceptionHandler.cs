using System;
using System.Configuration;

namespace FinancialAnalytics.Wrappers.Office.ExceptionHandling
{
    public static class ExceptionHandler
    {
        private static readonly object _locker = new object();
        private static IExceptionHandler _exceptionHandler;

        public static bool HandleException(Exception exception)
        {
            return GetHandler().HandleException(exception);
        }

        public static void LogException(Exception exception)
        {
            GetHandler().LogException(exception);
        }

        private static IExceptionHandler GetHandler()
        {
            lock (_locker)
            {
                if (_exceptionHandler != null)
                    return _exceptionHandler;

                object sectionObject = ConfigurationManager.GetSection("wrappersExceptionHandler");
                if (sectionObject == null || !(sectionObject is ExceptionHandlingConfigurationSettings))
                {
                    _exceptionHandler = new EmptyExceptionHandler();
                }
                else
                {
                    ExceptionHandlingConfigurationSettings section = (ExceptionHandlingConfigurationSettings)sectionObject;
                    Type handlerType = Type.GetType(section.HandlerType.TypeName);
                    if (handlerType == null)
                    {
                        _exceptionHandler = new EmptyExceptionHandler();
                        return _exceptionHandler;
                    }
                    object handlerObject = Activator.CreateInstance(handlerType);
                    if (handlerObject is IExceptionHandler)
                    {
                        _exceptionHandler = (IExceptionHandler)handlerObject;
                    }
                    else
                    {
                        _exceptionHandler = new EmptyExceptionHandler();
                    }
                }
                return _exceptionHandler;
            }
        }
    }
}
