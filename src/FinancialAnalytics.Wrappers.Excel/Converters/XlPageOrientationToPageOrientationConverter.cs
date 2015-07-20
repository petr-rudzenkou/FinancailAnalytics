using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public class XlPageOrientationToPageOrientationConverter
	{
		public static PageOrientation Convert(XlPageOrientation xlPageOrientation)
		{
			switch (xlPageOrientation)
			{
				case XlPageOrientation.xlLandscape:
					return PageOrientation.Landscape;
				case XlPageOrientation.xlPortrait:
					return PageOrientation.Portrait;
				default:
					return PageOrientation.Portrait;
			}
		}

		public static XlPageOrientation ConvertBack(PageOrientation pageOrientation)
		{
			switch (pageOrientation)
			{
				case PageOrientation.Landscape:
					return XlPageOrientation.xlLandscape;
				case PageOrientation.Portrait:
					return XlPageOrientation.xlPortrait;
				default:
					return XlPageOrientation.xlPortrait;
			}
		}
	}
}
