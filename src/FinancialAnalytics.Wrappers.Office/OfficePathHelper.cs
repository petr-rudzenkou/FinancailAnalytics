using System;
using Microsoft.Win32;

namespace FinancialAnalytics.Wrappers.Office
{
    public class OfficePathHelper
    {
        /// <summary>
        ///  Gets the Office program folder
        /// </summary>
        /// <returns></returns>
        public static string GetOfficePath(string applicationId)
        {
            string officeRegistryEntry = String.Format(@"SOFTWARE\Microsoft\Office\{0}.0\Common\InstallRoot", applicationId);
            string pathFromRegistry = GetValue(officeRegistryEntry, "Path").ToString();
            return pathFromRegistry;
        }

        public static object GetValue(string path, string keyName)
        {
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey(path);
            return regKey.GetValue(keyName);
        }
    }
}
