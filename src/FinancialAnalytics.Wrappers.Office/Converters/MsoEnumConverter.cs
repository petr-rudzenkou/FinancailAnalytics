using System;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public static class MsoEnumConverter
	{
		public static TWrapperResult ConvertFromMso<TMsoSource, TWrapperResult>(TMsoSource source)
		{
			string msoName = source.ToString();
			string name = msoName.Remove(0, 3);
			if (Enum.IsDefined(typeof(TWrapperResult), name))
			{
				return (TWrapperResult)Enum.Parse(typeof(TWrapperResult), name, true);
			}
			throw new ArgumentOutOfRangeException();
		}

		public static TMsoResult ConvertToMso<TWrapperSource, TMsoResult>(TWrapperSource source)
		{
			string name = source.ToString();
			string msoName = "mso" + name;
			if (Enum.IsDefined(typeof(TMsoResult), msoName))
			{
				return (TMsoResult)Enum.Parse(typeof(TMsoResult), msoName, true);
			}
			throw new ArgumentOutOfRangeException();
		}
	}
}