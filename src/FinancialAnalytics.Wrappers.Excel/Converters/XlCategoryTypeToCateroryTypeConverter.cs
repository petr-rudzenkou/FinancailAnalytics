using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{

	public class XlCategoryTypeToCateroryTypeConverter
	{
		public static CategoryType Convert(XlCategoryType xlCategoryType)
		{
			CategoryType categoryType;
			switch (xlCategoryType)
			{
				case XlCategoryType.xlCategoryScale:
					categoryType = CategoryType.CategoryScale;
					break;
				case XlCategoryType.xlTimeScale:
					categoryType = CategoryType.TimeScale;
					break;
				default:
					categoryType = CategoryType.AutomaticScale;
					break;
			}
			return categoryType;
		}

		public static XlCategoryType ConvertBack(CategoryType categoryType)
		{
			XlCategoryType xlCategoryType;
			switch (categoryType)
			{
				case CategoryType.TimeScale:
					xlCategoryType = XlCategoryType.xlTimeScale;
					break;
				case CategoryType.CategoryScale:
					xlCategoryType = XlCategoryType.xlCategoryScale;
					break;
				default:
					xlCategoryType = XlCategoryType.xlAutomaticScale;
					break;
			}
			return xlCategoryType;
		}
	}
}
