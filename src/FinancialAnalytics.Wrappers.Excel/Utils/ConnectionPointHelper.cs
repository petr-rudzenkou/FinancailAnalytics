using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Excel.Interception;
using FinancialAnalytics.Wrappers.Office;

namespace FinancialAnalytics.Wrappers.Excel.Utils
{
    public static class ConnectionPointHelper
    {
        private static string _processName;

        public static void SetupEventsConnection(Object unkSinc,
            IConnectionPointContainer connPointContainer, 
            string officeProcessName, 
            Guid eventsInterfaceGuid, 
            ref int cookie,
            out IConnectionPoint connectionPoint)
        {
            connectionPoint = null;
            using (new EnUsCultureInvoker())
            {
                if (string.IsNullOrEmpty(_processName))
                    _processName = Process.GetCurrentProcess().ProcessName;

                if (!_processName.Equals(officeProcessName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return;
                }
                if (cookie != 0)
                {
                    return;
                }
                connPointContainer.FindConnectionPoint(ref eventsInterfaceGuid, out connectionPoint);
                connectionPoint.Advise(unkSinc, out cookie);
            }
        }

        public static void RemoveEventsConnection(IConnectionPoint connectionPoint, bool isStarted, int cookie)
        {
            using (new EnUsCultureInvoker())
            {
                if (cookie != 0 && connectionPoint != null && isStarted)
                {
                    try
                    {
                        connectionPoint.Unadvise(cookie);
                    }
                    catch (Exception)
                    {
                    }
                    ComObjectsFinalizer.ReleaseComObject(connectionPoint);
                    connectionPoint = null;
                    cookie = 0;
                }
            }
        }
    }
}
