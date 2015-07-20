using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public static class XLSearchDirectionToSearchDirectionConverter
	{
		public static SearchDirection Convert(XlSearchDirection xlSearchDirection)
		{
			return (SearchDirection)(int)xlSearchDirection;
		}

		public static XlSearchDirection ConvertBack(SearchDirection searchDirection)
		{
			return (XlSearchDirection)(int)searchDirection;
		}
	}
}
