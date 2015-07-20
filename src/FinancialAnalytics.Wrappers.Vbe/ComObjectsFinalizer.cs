using System;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Vbe
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
                catch (Exception)
                {
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
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }
                obj = null;
            }
        }
    }
}
