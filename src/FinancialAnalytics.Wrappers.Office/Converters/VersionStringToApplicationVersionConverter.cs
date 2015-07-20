using System;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class VersionStringToApplicationVersionConverter
	{
		public static ApplicationVersion Convert(string versionStr)
		{
			if (string.IsNullOrEmpty(versionStr))
			{
				return ApplicationVersion.Undefined;
			}

			if (versionStr.IndexOf("11") >= 0)
			{
				return ApplicationVersion.Office2003;
			}
			else if (versionStr.IndexOf("12") >= 0)
			{
				return ApplicationVersion.Office2007;
			}
			else if (versionStr.IndexOf("14") >= 0)
			{
				return ApplicationVersion.Office2010;
			}
            else if (versionStr.IndexOf("15") >= 0)
            {
                return ApplicationVersion.Office2013;
            }
			else
			{
				return ApplicationVersion.Undefined;
			}
		}
	}
}
