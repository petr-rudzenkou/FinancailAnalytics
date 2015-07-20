using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlCmdTypeToCmdTypeConverter
    {
        public static CmdType Convert(XlCmdType xlCmdType)
        {
            CmdType result = CmdType.Default;
            switch (xlCmdType)
            {
                case XlCmdType.xlCmdCube:
                    result = CmdType.Cube;
                    break;
                case XlCmdType.xlCmdDefault:
                    result = CmdType.Default;
                    break;
                case XlCmdType.xlCmdList:
                    result = CmdType.List;
                    break;
                case XlCmdType.xlCmdSql:
                    result = CmdType.Sql;
                    break;
                case XlCmdType.xlCmdTable:
                    result = CmdType.Table;
                    break;
            }
            return result;
        }

        public static XlCmdType ConvertBack(CmdType cmdType)
        {
            XlCmdType result = XlCmdType.xlCmdDefault;
            switch (cmdType)
            {
                case CmdType.Cube:
                    result = XlCmdType.xlCmdCube;
                    break;
                case CmdType.Default:
                    result = XlCmdType.xlCmdDefault;
                    break;
                case CmdType.List:
                    result = XlCmdType.xlCmdList;
                    break;
                case CmdType.Sql:
                    result = XlCmdType.xlCmdSql;
                    break;
                case CmdType.Table:
                    result = XlCmdType.xlCmdTable;
                    break;
            }
            return result;
        }
    }
}
