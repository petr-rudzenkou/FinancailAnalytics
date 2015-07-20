using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlPivotTableVersionListToPivotTableVersionListConverter
    {
        public static PivotTableVersionList Convert(XlPivotTableVersionList xlVersion)
        {
            return (PivotTableVersionList) (int) xlVersion;
        }

        public static XlPivotTableVersionList ConvertBack(PivotTableVersionList version)
        {
            XlPivotTableVersionList result = XlPivotTableVersionList.xlPivotTableVersionCurrent;
            switch (version)
            {
                case PivotTableVersionList.Version10:
                    result = XlPivotTableVersionList.xlPivotTableVersion10;
                    break;
                case PivotTableVersionList.Version2000:
                    result = XlPivotTableVersionList.xlPivotTableVersion2000;
                    break;
                case PivotTableVersionList.VersionCurrent:
                    result = XlPivotTableVersionList.xlPivotTableVersionCurrent;
                    break;
            }
            return result;
        }
    }
}
