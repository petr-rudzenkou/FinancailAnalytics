using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlSaveAsAccessModeToSaveAsAccessModeConverter
    {
        public static SaveAsAccessMode Convert(XlSaveAsAccessMode xlSaveAsAccessMode)
        {
            SaveAsAccessMode result = SaveAsAccessMode.Shared;
            switch (xlSaveAsAccessMode)
            {
                case XlSaveAsAccessMode.xlExclusive:
                    result = SaveAsAccessMode.Exclusive;
                    break;
                case XlSaveAsAccessMode.xlNoChange:
                    result = SaveAsAccessMode.NoChange;
                    break;
                case XlSaveAsAccessMode.xlShared:
                    result = SaveAsAccessMode.Shared;
                    break;
            }
            return result;
        }

        public static XlSaveAsAccessMode ConvertBack(SaveAsAccessMode saveAsAccessMode)
        {
            XlSaveAsAccessMode result = XlSaveAsAccessMode.xlShared;
            switch (saveAsAccessMode)
            {
                case SaveAsAccessMode.Exclusive:
                    result = XlSaveAsAccessMode.xlExclusive;
                    break;
                case SaveAsAccessMode.NoChange:
                    result = XlSaveAsAccessMode.xlNoChange;
                    break;
                case SaveAsAccessMode.Shared:
                    result = XlSaveAsAccessMode.xlShared;
                    break;
            }
            return result;
        }
    }
}
