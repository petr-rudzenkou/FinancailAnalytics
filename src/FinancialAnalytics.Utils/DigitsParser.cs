using System;
using System.Globalization;

namespace FinancialAnalytics.Utils
{
    public class DigitsParser
    {
        public static double GetDoubleForMarCap(string value, double defaultValue)
        {
            double result;
            char cap = '_';
            bool lastSymbolIsNotDigit = !Char.IsDigit(value[value.Length - 1]);
            if (lastSymbolIsNotDigit)
            {
                cap = value[value.Length - 1];
                value = value.Substring(0, value.Length - 1);
            }

            //Try parsing in the current culture
            if (!double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out result) &&
                //Then try in US english
                !double.TryParse(value, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result) &&
                //Then in neutral language
                !double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result = defaultValue;
            }

            switch (cap)
            {
                case '_':
                    result = result / 1000000; 
                    break;
                case 'B':
                    result = result * 1000;
                    break;
                case 'M':
                    break;
            }

            return result;
        }

        public static double GetDouble(string value, double defaultValue)
        {
            double result;

            //Try parsing in the current culture
            if (!double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out result) &&
                //Then try in US english
                !double.TryParse(value, NumberStyles.Any, CultureInfo.GetCultureInfo("en-US"), out result) &&
                //Then in neutral language
                !double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result = defaultValue;
            }

            
            return result;
        }

        private static int FindFirstNonDigit(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                if (!(Char.IsDigit(s[i]))) return i;
            }
            return -1;
        }
    }
}
