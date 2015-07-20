using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoShadowTypeToShadowTypeConverter
	{
		public static ShadowType Convert(int msoControlType)
		{			
			return (ShadowType)msoControlType;		
		}

		public static int ConvertBack(ShadowType controlType)
		{
			return (int)controlType;
		}
	}
}
