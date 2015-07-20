using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoGradientStyleToGradientStyleConverter
	{
		public static GradientStyle Convert(MsoGradientStyle msoGradientStyle)
		{
			return MsoEnumConverter.ConvertFromMso<MsoGradientStyle, GradientStyle>(msoGradientStyle);
		}

		public static MsoGradientStyle ConvertBack(GradientStyle gradientStyle)
		{
			return MsoEnumConverter.ConvertToMso<GradientStyle, MsoGradientStyle>(gradientStyle);
		}
	}
}
