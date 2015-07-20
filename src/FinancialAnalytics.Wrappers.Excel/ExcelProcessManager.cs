using System;
using System.Linq;
using System.Collections.Generic;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.Windows;
using NativeMethods = FinancialAnalytics.Wrappers.Office.Windows.NativeMethods;

namespace FinancialAnalytics.Wrappers.Excel
{
    internal class ExcelProcessManager
    {
		private const string ExcelApplicationProgId = "Excel.Application";

        public IDictionary<Microsoft.Office.Interop.Excel.Application, uint> GetApplications()
        {
        	return GetApplications(false);
        }

        public IDictionary<Microsoft.Office.Interop.Excel.Application, uint> GetApplications(bool includeHidden)
        {
            Dictionary<Microsoft.Office.Interop.Excel.Application, uint> result = new Dictionary<Microsoft.Office.Interop.Excel.Application, uint>();

            foreach (var hwnd in OfficeWindowsManager.FindExcelWindows())
            {
                if (!NativeMethods.IsWindowVisible(hwnd) && !includeHidden) // it's invisible instance (embedded)
                    continue;

                Microsoft.Office.Interop.Excel.Application application = GetApplication(hwnd);
                if (application != null)
                {
                    if (!application.UserControl)
                    {
						ComObjectsFinalizer.ReleaseComObject(application); 
                        continue;
                    }
					
					//When running from office 2013, OfficeWindowManager returns several Excel window handles, that point to the same Application instances.
					//To prevent failure, when adding entry with the same key to Dictinary<TKey, TValue>, next two rows check key's existence before that.
					if (result.ContainsKey(application))
						continue;

                    uint processId = (uint)NativeWindowsManager.GetProcessId(hwnd);
                    result.Add(application, processId);
                }
            }

            return result;            
        }

        public Microsoft.Office.Interop.Excel.Application GetApplication(IntPtr mainWindowHandler)
        {
            var window = (Microsoft.Office.Interop.Excel.Window)OfficeWindowsManager.FindExcelWindowObject(mainWindowHandler);
            if( window == null)
			{
				object activeObject = null;
				try
				{
					activeObject = System.Runtime.InteropServices.Marshal.GetActiveObject(ExcelApplicationProgId);
				}
				catch(Exception)
				{
					// If Excel is not running, GetActiveObject() will throw an exception
					return null;
				}

				if (activeObject != null)
				{
					Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)activeObject;
					if (application != null)
					{
						if (application.Hwnd != mainWindowHandler.ToInt32())
						{
							ComObjectsFinalizer.ReleaseComObject(application);
							return null;
						}
						return application;
					}
					ComObjectsFinalizer.ReleaseComObject(activeObject);
				}
			}
		    return window == null ? null : window.Application;
        }

        public Microsoft.Office.Interop.Excel.Application GetActive(string version)
        {
        	return GetActive(version, false);
        }

		public Microsoft.Office.Interop.Excel.Application GetActive(string version, bool includeHidden)
		{
			return GetApplications(includeHidden).Select(x => x.Key).FirstOrDefault();
		}
        
    }
}
