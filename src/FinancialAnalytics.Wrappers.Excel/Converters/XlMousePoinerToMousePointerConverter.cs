using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlMousePoinerToMousePointerConverter
    {
        public MousePointer Convert(XlMousePointer xlMousePoiner)
        {
            MousePointer result;
            switch (xlMousePoiner)
            {
                case XlMousePointer.xlDefault:
                    result = MousePointer.Beam;
                    break;
                case XlMousePointer.xlIBeam:
                    result = MousePointer.Beam;
                    break;
                case XlMousePointer.xlNorthwestArrow:
                    result = MousePointer.NorthwestArrow;
                    break;
                case XlMousePointer.xlWait:
                    result = MousePointer.Wait;
                    break;
                default:
                    throw new InvalidEnumArgumentException("xlMousePoiner");
            }
            return result;
        }

        public XlMousePointer ConvertBack(MousePointer mousePointer)
        {
            XlMousePointer result;
            switch (mousePointer)
            {
                case MousePointer.Beam:
                    result = XlMousePointer.xlIBeam;
                    break;
                case MousePointer.Default:
                    result = XlMousePointer.xlDefault;
                    break;
                case MousePointer.NorthwestArrow:
                    result = XlMousePointer.xlNorthwestArrow;
                    break;
                case MousePointer.Wait:
                    result = XlMousePointer.xlWait;
                    break;
                default:
                    throw new InvalidEnumArgumentException("mousePointer");
            }
            return result;
        }
    }
}
