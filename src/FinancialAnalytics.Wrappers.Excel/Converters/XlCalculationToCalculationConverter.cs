using Microsoft.Office.Interop.Excel;
using FinancialAnalytics.Wrappers.Excel.Enums;

namespace FinancialAnalytics.Wrappers.Excel.Converters
{
    public class XlCalculationToCalculationConverter
    {
        public static Calculation Convert(XlCalculation xlCalculation)
        {
            Calculation calculation;
            switch (xlCalculation)
            {
                case XlCalculation.xlCalculationSemiautomatic:
                    calculation = Calculation.Semiautomatic;
                    break;
                case XlCalculation.xlCalculationManual:
                    calculation = Calculation.Manual;
                    break;
                default:
                    calculation = Calculation.Automatic;
                    break;
            }
            return calculation;
        }

        public static XlCalculation ConvertBack(Calculation calculation)
        {
            XlCalculation xlCalculation;
            switch (calculation)
            {
                case Calculation.Semiautomatic:
                    xlCalculation = XlCalculation.xlCalculationSemiautomatic;
                    break;
                case Calculation.Manual:
                    xlCalculation = XlCalculation.xlCalculationManual;
                    break;
                default:
                    xlCalculation = XlCalculation.xlCalculationAutomatic;
                    break;
            }
            return xlCalculation;
        }
    }
}
