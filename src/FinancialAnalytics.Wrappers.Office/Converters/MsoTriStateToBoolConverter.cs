using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FinancialAnalytics.Wrappers.Office.Enums;
using Microsoft.Office.Core;

namespace FinancialAnalytics.Wrappers.Office.Converters
{
    public class MsoTriStateToBoolConverter
    {
        public static bool Convert(MsoTriState msoTriState)
        {
            bool value = false;
            if (msoTriState == MsoTriState.msoTrue)
            {
                value = true;
            }
            return value;
        }

        public static MsoTriState ConvertBack(bool value)
        {
            MsoTriState msoTriState = MsoTriState.msoFalse;
            if (value)
            {
                msoTriState = MsoTriState.msoTrue;
            }
            return msoTriState;
        }
    }
}
