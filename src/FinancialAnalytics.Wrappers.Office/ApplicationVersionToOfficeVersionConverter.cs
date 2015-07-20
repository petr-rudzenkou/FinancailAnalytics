using System;
using System.Reflection;
using FinancialAnalytics.Wrappers.Office;
using FinancialAnalytics.Wrappers.Office.ExceptionHandling;

namespace FinancialAnalytics.Wrappers.Office
{

    public class ApplicationVersionToOfficeVersionConverter
    {
        public const string Office2003ApplicationVersion = "11.0";
        public const string Office2007ApplicationVersion = "12.0";
        public const string Office2010ApplicationVersion = "14.0";
        public const string Office2013ApplicationVersion = "15.0";

        public static OfficeVersion Convert(string applicationVersion)
        {
            OfficeVersion version = OfficeVersion.Other;
            switch (applicationVersion)
            {
                case Office2003ApplicationVersion:
                    version = OfficeVersion.Office2003;
                    break;
                case Office2007ApplicationVersion:
                    version = OfficeVersion.Office2007;
                    break;
                case Office2010ApplicationVersion:
                    version = OfficeVersion.Office2010;
                    break;
                case Office2013ApplicationVersion:
                    version = OfficeVersion.Office2013;
                    break;
            }
            return version;
        }

        public static OfficeVersion Convert(object application)
        {
            try
            {
				string applicationVersion = null;
					applicationVersion = (string)application.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, application, null);
                return Convert(applicationVersion);
            }
            catch (Exception exc)
            {
                bool rethrow = ExceptionHandler.HandleException(exc);
                if (rethrow)
                    throw;
                return OfficeVersion.Other;
            }
        }
    }
}
