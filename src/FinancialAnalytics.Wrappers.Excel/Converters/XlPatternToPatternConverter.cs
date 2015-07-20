using System;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public class XlPatternToPatternConverter
	{
		public static Pattern Convert(XlPattern xlPattern)
		{
			if ((int)xlPattern == (int)Pattern.PatternLinearGradient)
			{
				return Pattern.PatternLinearGradient;
			}
			if ((int)xlPattern == (int)Pattern.PatternRectangularGradient)
			{
				return Pattern.PatternRectangularGradient;
			}

			switch (xlPattern)
			{
				case XlPattern.xlPatternAutomatic:
					return Pattern.PatternAutomatic;
				case XlPattern.xlPatternUp:
					return Pattern.PatternUp;
				case XlPattern.xlPatternNone:
					return Pattern.PatternNone;
				case XlPattern.xlPatternHorizontal:
					return Pattern.PatternHorizontal;
				case XlPattern.xlPatternGray75:
					return Pattern.PatternGray75;
				case XlPattern.xlPatternGray50:
					return Pattern.PatternGray50;
				case XlPattern.xlPatternGray25:
					return Pattern.PatternGray25;
				case XlPattern.xlPatternDown:
					return Pattern.PatternDown;
				case XlPattern.xlPatternSolid:
					return Pattern.PatternSolid;
				case XlPattern.xlPatternChecker:
					return Pattern.PatternChecker;
				case XlPattern.xlPatternSemiGray75:
					return Pattern.PatternSemiGray75;
				case XlPattern.xlPatternLightHorizontal:
					return Pattern.PatternLightHorizontal;
				case XlPattern.xlPatternLightVertical:
					return Pattern.PatternLightVertical;
				case XlPattern.xlPatternLightDown:
					return Pattern.PatternLightDown;
				case XlPattern.xlPatternLightUp:
					return Pattern.PatternLightUp;
				case XlPattern.xlPatternGrid:
					return Pattern.PatternGrid;
				case XlPattern.xlPatternCrissCross:
					return Pattern.PatternCrissCross;
				case XlPattern.xlPatternGray16:
					return Pattern.PatternGray16;
				case XlPattern.xlPatternGray8:
					return Pattern.PatternGray8;
				default:
					return Pattern.PatternNone;
			}
		}

		public static XlPattern ConvertBack(Pattern pattern)
		{
			if (pattern == Pattern.PatternLinearGradient)
			{
				return (XlPattern)Pattern.PatternLinearGradient;
			}
			if (pattern == Pattern.PatternRectangularGradient)
			{
				return (XlPattern)Pattern.PatternRectangularGradient;
			}

			switch (pattern)
			{
				case Pattern.PatternAutomatic:
					return XlPattern.xlPatternAutomatic;
				case Pattern.PatternUp:
					return XlPattern.xlPatternUp;
				case Pattern.PatternNone:
					return XlPattern.xlPatternNone;
				case Pattern.PatternHorizontal:
					return XlPattern.xlPatternHorizontal;
				case Pattern.PatternGray75:
					return XlPattern.xlPatternGray75;
				case Pattern.PatternGray50:
					return XlPattern.xlPatternGray50;
				case Pattern.PatternGray25:
					return XlPattern.xlPatternGray25;
				case Pattern.PatternDown:
					return XlPattern.xlPatternDown;
				case Pattern.PatternSolid:
					return XlPattern.xlPatternSolid;
				case Pattern.PatternChecker:
					return XlPattern.xlPatternChecker;
				case Pattern.PatternSemiGray75:
					return XlPattern.xlPatternSemiGray75;
				case Pattern.PatternLightHorizontal:
					return XlPattern.xlPatternLightHorizontal;
				case Pattern.PatternLightVertical:
					return XlPattern.xlPatternLightVertical;
				case Pattern.PatternLightDown:
					return XlPattern.xlPatternLightDown;
				case Pattern.PatternLightUp:
					return XlPattern.xlPatternLightUp;
				case Pattern.PatternGrid:
					return XlPattern.xlPatternGrid;
				case Pattern.PatternCrissCross:
					return XlPattern.xlPatternCrissCross;
				case Pattern.PatternGray16:
					return XlPattern.xlPatternGray16;
				case Pattern.PatternGray8:
					return XlPattern.xlPatternGray8;
				default:
					return XlPattern.xlPatternNone;
			}
		}

	}
}
