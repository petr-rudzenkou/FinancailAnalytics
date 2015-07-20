using System;
using System.Runtime.InteropServices;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;

namespace FinancialAnalytics.Wrappers.Office
{
    public class ComObjectsFinalizer
    {
        public static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception exc)
                {
                    bool rethrow = ExceptionHandler.HandleException(exc);
                    if (rethrow)
                        throw;
                }
                finally
                {
                    obj = null;
                }
            }
        }

        public static void FinalReleaseComObject(object obj)
        {
            if (obj != null)
            {
                while (true)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(obj);
                        break;
                    }
                    catch (Exception exc)
                    {
                        System.Threading.Thread.Sleep(100);
                        bool rethrow = ExceptionHandler.HandleException(exc);
                        if (rethrow)
                            throw;
                    }
                }
                obj = null;
            }
        }
    }
}
