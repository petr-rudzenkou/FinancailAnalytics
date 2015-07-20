using Microsoft.Office.Core;
using FinancialAnalytics.Wrappers.Office.Enums;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
	public class MsoFileDialogTypeToFileDialogTypeConverter
	{
		public static FileDialogType Convert(MsoFileDialogType msoFileDialogType)
		{
			return MsoEnumConverter.ConvertFromMso<MsoFileDialogType, FileDialogType>(msoFileDialogType);
		}

		public static MsoFileDialogType ConvertBack(FileDialogType fileDialogType)
		{
			return MsoEnumConverter.ConvertToMso<FileDialogType, MsoFileDialogType>(fileDialogType);
		}
	}
}
