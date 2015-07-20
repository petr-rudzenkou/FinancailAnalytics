using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office.Windows
{
    public static class TopWindowsFinder
    {
        public static IEnumerable<IntPtr> Find(string className, string windowName = null)
        {
            IntPtr handle = IntPtr.Zero;
            while (true)
            {
                handle = NativeMethods.FindWindowEx(IntPtr.Zero, handle, className, windowName);
                if (handle == IntPtr.Zero)
                {
                    yield break;
                }
                else
                {
                    yield return handle;
                }
            }
        }
    }
}