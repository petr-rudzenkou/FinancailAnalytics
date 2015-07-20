using System;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
	public static class XlCellTypeToCellTypeConverter
	{
		public static CellType Convert(XlCellType xlCellType)
		{
			CellType cellType = CellType.AllFormatConditions;
			string cellTypeName = xlCellType.ToString();
			cellTypeName = cellTypeName.Remove(0, 10);
			if (Enum.IsDefined(typeof(CellType), cellTypeName))
			{
				cellType = (CellType)Enum.Parse(typeof(CellType), cellTypeName, true);
			}
			return cellType;
		}

		public static XlCellType Convert(CellType cellType)
		{
			XlCellType xlCellType = XlCellType.xlCellTypeAllFormatConditions;
			string cellTypeName = cellType.ToString();
			cellTypeName = cellTypeName.Insert(0, "xlCellType");
			if (Enum.IsDefined(typeof(XlCellType), cellTypeName))
			{
				xlCellType = (XlCellType)Enum.Parse(typeof(XlCellType), cellTypeName, true);
			}
			return xlCellType;
		}
	}
}
