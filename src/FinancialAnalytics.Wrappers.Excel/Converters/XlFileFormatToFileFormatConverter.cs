using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public static class XlFileFormatToFileFormatConverter
    {
        public static FileFormat Convert(XlFileFormat xlFileFormat)
        {
            return (FileFormat) (int) xlFileFormat;
        }

        public static XlFileFormat ConvertBack(FileFormat fileFormat)
        {
            return (XlFileFormat) (int) fileFormat;
        }
    }
}
